;;; ============================================================
;;; findreplace.lsp
;;; AutoCAD / Civil3D Find & Replace for TEXT and MTEXT entities
;;;
;;; Load via APPLOAD (or drag-drop into the drawing window),
;;; then type FINDREPLACE at the Command prompt.
;;;
;;; Workflow
;;;   1. Prompt for a search string
;;;   2. Collect every TEXT / MTEXT entity whose content
;;;      contains the search string (case-insensitive)
;;;   3. Highlight all matches in yellow (color index 2)
;;;   4. Prompt for a replacement string
;;;   5. Step through each match one at a time:
;;;        Zoom in, display current text, ask Replace / Keep / ESC
;;; ============================================================


;;; ---- String helpers ----------------------------------------

;;; Returns T if NEEDLE appears anywhere in HAYSTACK
;;; (case-insensitive)
(defun fr:strfind (needle haystack)
  (not (null (vl-string-search (strcase needle)
                               (strcase haystack))))
)

;;; Replaces every case-insensitive occurrence of NEEDLE with
;;; REPLACEMENT inside STR, preserving the original casing of
;;; surrounding characters (important for MTEXT formatting codes).
(defun fr:strreplaceall (needle replacement str
                         / nlen pos result rest)
  (setq nlen   (strlen needle)
        result ""
        rest   str)
  (while (setq pos (vl-string-search (strcase needle)
                                     (strcase rest)))
    ;; Append the unchanged prefix, then the replacement text.
    ;; (substr rest 1 pos) pulls from the original, preserving case.
    (setq result (strcat result (substr rest 1 pos) replacement)
          rest   (substr rest (+ pos nlen 1)))
  )
  (strcat result rest)    ; append any trailing characters
)


;;; ---- Entity helpers ----------------------------------------

;;; Return the raw text string of a TEXT or MTEXT entity.
;;; For MTEXT this includes inline formatting codes.
(defun fr:gettext (ent)
  (cdr (assoc 1 (entget ent)))
)

;;; Return the entity's current color index.
;;; Returns 256 (BYLAYER) when no explicit color is stored.
(defun fr:getcolor (ent / pair)
  (setq pair (assoc 62 (entget ent)))
  (if pair (cdr pair) 256)
)

;;; Set COLOR-VAL on ENT.  Pass 256 to restore BYLAYER
;;; (removes group 62 so the entity inherits the layer color).
(defun fr:setcolor (ent colorval / dxf)
  (setq dxf (entget ent)
        dxf (vl-remove (assoc 62 dxf) dxf))   ; strip existing color
  (if (/= colorval 256)
    (setq dxf (append dxf (list (cons 62 colorval))))
  )
  (entmod dxf)
  (entupd ent)
)

;;; Write NEWSTR back into ENT (works for both TEXT and MTEXT).
(defun fr:settext (ent newstr / dxf)
  (setq dxf (entget ent))
  (entmod (subst (cons 1 newstr) (assoc 1 dxf) dxf))
  (entupd ent)
)


;;; ---- Viewport helper ---------------------------------------

;;; Zoom to ENT with generous padding around its bounding box.
;;; Falls back to ZOOM Object when getboundingbox is unavailable.
(defun fr:zooment (ent / obj minpt maxpt dx dy pad ss)
  (setq obj (vlax-ename->vla-object ent))
  (if (not (vl-catch-all-error-p
              (vl-catch-all-apply 'vla-getboundingbox
                                  (list obj 'minpt 'maxpt))))
    ;; Bounding box succeeded – build a padded window
    (progn
      (setq minpt (vlax-safearray->list minpt)
            maxpt (vlax-safearray->list maxpt)
            dx    (- (car  maxpt) (car  minpt))
            dy    (- (cadr maxpt) (cadr minpt))
            ;; Padding = 75 % of the longer dimension, minimum 0.1
            pad   (* (max dx dy 0.1) 0.75))
      (command "_.ZOOM" "_W"
               (list (- (car  minpt) pad) (- (cadr minpt) pad))
               (list (+ (car  maxpt) pad) (+ (cadr maxpt) pad)))
    )
    ;; Fallback – ZOOM Object on a one-item selection set
    (progn
      (setq ss (ssadd ent (ssadd)))
      (command "_.ZOOM" "_O" ss "")
    )
  )
)


;;; ---- Color restore helper ----------------------------------

;;; Restore original colors for every entity in MATCH-LIST
;;; whose index is >= START-IDX.
(defun fr:restorecolors (match-list colors start-idx / i)
  (setq i start-idx)
  (while (< i (length match-list))
    (fr:setcolor (nth i match-list) (nth i colors))
    (setq i (1+ i))
  )
)


;;; ---- Main command ------------------------------------------

(defun c:FINDREPLACE
       (/ searchstr replstr ss slen idx ent
          colors match-list total choice
          textval newtext done)

  (vl-load-com)
  (princ "\n=== FINDREPLACE ===")

  ;; ----------------------------------------------------------
  ;; Step 1 – Search string
  ;; T as first argument to getstring permits embedded spaces.
  ;; ----------------------------------------------------------
  (setq searchstr (getstring T "\nEnter search string: "))

  (if (= searchstr "")

    ;; Nothing entered – abort cleanly
    (princ "\nNo search string entered. Exiting.")

    ;; --------------------------------------------------------
    ;; Main body (wrapped so we can skip it on empty input)
    ;; --------------------------------------------------------
    (progn

      ;; -------------------------------------------------------
      ;; Step 2 – Collect matching entities from ALL spaces
      ;; "_X" flag = search entire database (model + all layouts)
      ;; -------------------------------------------------------
      (setq match-list '())
      (setq ss (ssget "_X" '((0 . "TEXT,MTEXT"))))

      (if ss
        (progn
          (setq slen (sslength ss)
                idx  0)
          (repeat slen
            (setq ent     (ssname ss idx)
                  textval (fr:gettext ent))
            (if (and textval (fr:strfind searchstr textval))
              (setq match-list (cons ent match-list))
            )
            (setq idx (1+ idx))
          )
          (setq match-list (reverse match-list)) ; preserve draw order
        )
      )

      (setq total (length match-list))

      ;; -------------------------------------------------------
      ;; No matches found
      ;; -------------------------------------------------------
      (if (= total 0)

        (princ (strcat "\nNo matches found for: \""
                       searchstr "\". Exiting."))

        ;; -----------------------------------------------------
        ;; Matches found – highlight, collect originals
        ;; -----------------------------------------------------
        (progn

          ;; Save original colors then paint all matches yellow
          (setq colors '())
          (foreach ent match-list
            (setq colors (append colors (list (fr:getcolor ent))))
            (fr:setcolor ent 2)   ; 2 = yellow
          )

          (princ (strcat "\nFound "
                         (itoa total)
                         " match(es). Press Enter to continue..."))
          (getstring "")           ; pause for user to inspect

          ;; ---------------------------------------------------
          ;; Step 3 – Replacement string
          ;; ---------------------------------------------------
          (setq replstr (getstring T "\nEnter replacement string: "))

          ;; ---------------------------------------------------
          ;; Step 4 – Review loop (one entity at a time)
          ;; ---------------------------------------------------
          (setq idx  0
                done nil)

          (while (and (< idx total) (not done))

            (setq ent     (nth idx match-list)
                  textval (fr:gettext ent))

            ;; Zoom to the current entity
            (fr:zooment ent)

            ;; Show position counter and current text content
            (princ (strcat "\n[" (itoa (1+ idx)) "/"
                           (itoa total) "]  Text: \""
                           textval "\""))

            ;; getkword presents R/K; nil is returned on ESC
            (initget "Replace Keep")
            (setq choice
              (getkword
                "\n  Action [Replace/Keep] or ESC to cancel: "))

            (cond

              ;; ----------- Replace --------------------------
              ((= choice "Replace")
               (setq newtext (fr:strreplaceall
                               searchstr replstr textval))
               (fr:settext   ent newtext)
               (fr:setcolor  ent (nth idx colors)) ; restore color
               (setq idx (1+ idx))
              )

              ;; ----------- Keep -----------------------------
              ((= choice "Keep")
               (fr:setcolor ent (nth idx colors))  ; restore color
               (setq idx (1+ idx))
              )

              ;; ----------- ESC / nil ------------------------
              (T
               ;; Restore highlights for this and all remaining
               (fr:restorecolors match-list colors idx)
               (princ "\nCancelled. All remaining highlights restored.")
               (setq done T)
              )
            ) ; end cond
          ) ; end while

          (if (not done)
            (princ "\nFind & Replace complete.")
          )
        ) ; end progn (matches found)
      ) ; end if total = 0
    ) ; end progn (main body)
  ) ; end if searchstr empty

  (princ)   ; suppress the nil echo at the Command prompt
)

(princ "\nfindreplace.lsp loaded.  Type FINDREPLACE to run.")
(princ)
;;; ============================================================
;;; End of findreplace.lsp
;;; ============================================================
