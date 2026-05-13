;;; ============================================================
;;; findreplace.lsp  (v2.1)
;;; AutoCAD / Civil3D — Find & Replace across drawing text entities
;;;
;;; Supported entity types
;;;   TEXT        plain single-line text             (DXF 1)
;;;   MTEXT       multiline / formatted text         (DXF 1)
;;;   DIMENSION   user override text only            (DXF 1; skips "<>" / empty)
;;;   ATTRIB      attribute value in a block ref     (DXF 1)
;;;   ATTDEF      attribute definition (template)    (DXF 1)
;;;   MULTILEADER MLeader carrying MTEXT content     (ActiveX TextString)
;;;   LEADER      annotation text is a separate MTEXT/TEXT entity and is
;;;               therefore picked up by those filters automatically
;;;
;;; Civil3D label-style annotations (AECC_* proxy entities) use a
;;; different DXF entity name and are excluded by the ssget filter.
;;; Change those through the Label Style Composer instead.
;;;
;;; Highlighting: AutoCAD's native selection glow (redraw modes 3 / 4).
;;; No entity color properties are ever read or written, so no color
;;; save/restore is required and no colors are left changed if the
;;; session is interrupted.
;;;
;;; Progress: each replacement is committed immediately when you choose
;;; Replace.  Pressing ESC at any prompt keeps every replacement made
;;; so far — nothing is rolled back.
;;;
;;; MTEXT width: after a replacement the text-box width is auto-expanded
;;; to the content's natural (un-wrapped) width, preventing the longer
;;; replacement word from forcing the next word onto a new line.
;;; Hard paragraph breaks (\P) are respected; only soft wrap is removed.
;;;
;;; Load via APPLOAD (or drag-drop into the drawing window),
;;; then type FINDREPLACE at the Command prompt.
;;; ============================================================

(vl-load-com)


;;; ════════════════════════════════════════════════════════════
;;; String utilities
;;; ════════════════════════════════════════════════════════════

(defun fr:strfind (needle haystack)
  ;; T when NEEDLE appears anywhere in HAYSTACK, case-insensitive
  (not (null (vl-string-search (strcase needle) (strcase haystack))))
)

(defun fr:strreplaceall (needle replacement str / nlen pos result rest)
  ;; Replace every case-insensitive occurrence of NEEDLE with REPLACEMENT.
  ;; The unchanged prefix before each match is taken from the original STR,
  ;; so MTEXT inline formatting codes that bracket the match are preserved.
  (setq nlen   (strlen needle)
        result ""
        rest   str)
  (while (setq pos (vl-string-search (strcase needle) (strcase rest)))
    (setq result (strcat result (substr rest 1 pos) replacement)
          rest   (substr rest (+ pos nlen 1)))
  )
  (strcat result rest)
)


;;; ════════════════════════════════════════════════════════════
;;; Entity text — get / set
;;; ════════════════════════════════════════════════════════════

(defun fr:gettext (ent / etype txt obj res)
  ;; Return the searchable text content of ENT, or nil when the entity
  ;; carries no user-editable text (DIMENSION with auto-measurement,
  ;; MLeader with block content, empty MLeader, etc.).
  (setq etype (cdr (assoc 0 (entget ent))))
  (cond
    ;; Plain DXF-group-1 entities
    ((member etype '("TEXT" "MTEXT" "ATTRIB" "ATTDEF"))
     (cdr (assoc 1 (entget ent))))

    ;; DIMENSION: only when the user has supplied an override string.
    ;; An empty string or "<>" means "show the measured value" — skip it.
    ((= etype "DIMENSION")
     (setq txt (cdr (assoc 1 (entget ent))))
     (if (and txt (/= txt "") (/= txt "<>")) txt nil))

    ;; MULTILEADER: access via ActiveX so we don't have to parse the
    ;; deeply-nested MLEADERCONTEXT DXF structure.
    ;; Returns nil for block-content leaders and on any ActiveX error.
    ((= etype "MULTILEADER")
     (setq obj (vlax-ename->vla-object ent)
           res (vl-catch-all-apply 'vla-get-textstring (list obj)))
     (if (or (vl-catch-all-error-p res) (= res "")) nil res))

    (T nil)
  )
)

(defun fr:settext (ent newstr / etype dxf obj)
  ;; Write NEWSTR back into ENT (handles all supported types).
  (setq etype (cdr (assoc 0 (entget ent))))
  (cond
    ((member etype '("TEXT" "MTEXT" "ATTRIB" "ATTDEF" "DIMENSION"))
     (setq dxf (entget ent))
     (entmod (subst (cons 1 newstr) (assoc 1 dxf) dxf))
     (entupd ent))
    ((= etype "MULTILEADER")
     (setq obj (vlax-ename->vla-object ent))
     (vl-catch-all-apply 'vla-put-textstring (list obj newstr))
     (entupd ent))
  )
)


;;; ════════════════════════════════════════════════════════════
;;; Visual highlight — selection glow, zero color-property changes
;;; ════════════════════════════════════════════════════════════

(defun fr:highlight   (ent) (redraw ent 3))   ; blue/white selection glow
(defun fr:dehighlight (ent) (redraw ent 4))   ; remove glow

(defun fr:dehighlightfrom (lst start-idx / i)
  ;; Remove glow from every entity in LST at position >= START-IDX.
  (setq i start-idx)
  (while (< i (length lst))
    (fr:dehighlight (nth i lst))
    (setq i (1+ i))
  )
)


;;; ════════════════════════════════════════════════════════════
;;; MTEXT width auto-adjustment
;;; ════════════════════════════════════════════════════════════

(defun fr:fixmtextwidth (ent / dxf widpair curw obj minpt maxpt natw)
  ;; Widen the MTEXT box to the content's natural (un-wrapped) width.
  ;; Steps:
  ;;   1. Temporarily set DXF 41 (box width) to 0 — no constraint
  ;;   2. Force AutoCAD to re-flow the text layout via vla-regen
  ;;      (entupd alone queues the update but the geometry cache used by
  ;;      vla-getboundingbox is not refreshed until the regen runs)
  ;;   3. Measure the actual rendered width via vla-getboundingbox
  ;;   4. Lock that natural width back into DXF 41
  ;; This eliminates soft wrap caused by the replacement being longer
  ;; than the original.  Hard \P paragraph breaks are unaffected.
  ;; Returns T so the caller knows a regen was performed and can
  ;; re-apply selection highlights that regen clears.
  (setq dxf     (entget ent)
        widpair (assoc 41 dxf)
        curw    (if widpair (cdr widpair) 0.0))
  (if (> curw 0.0)
    (progn
      ;; Release width constraint and force a re-layout
      (entmod (subst (cons 41 0.0) widpair dxf))
      (entupd ent)
      ;; acActiveViewport = 1; regen flushes the deferred layout cache
      (vla-regen (vla-get-activedocument (vlax-get-acad-object)) 1)
      (setq obj (vlax-ename->vla-object ent))
      (if (not (vl-catch-all-error-p
                  (vl-catch-all-apply 'vla-getboundingbox
                                      (list obj 'minpt 'maxpt))))
        ;; Bounding box succeeded — set natural width
        (progn
          (setq natw (- (car (vlax-safearray->list maxpt))
                        (car (vlax-safearray->list minpt))))
          (entmod (subst (cons 41 natw)
                         (assoc 41 (entget ent))
                         (entget ent)))
          (entupd ent)
        )
        ;; Bounding box call failed — restore the original width
        (progn
          (entmod (subst (cons 41 curw)
                         (assoc 41 (entget ent))
                         (entget ent)))
          (entupd ent)
        )
      )
      T   ; signal to caller: regen was performed, highlights were cleared
    )
  )
)


;;; ════════════════════════════════════════════════════════════
;;; Viewport zoom to a single entity
;;; ════════════════════════════════════════════════════════════

(defun fr:zooment (ent / obj minpt maxpt dx dy pad ss)
  ;; Zoom to ENT with generous padding.  Falls back to ZOOM Object
  ;; if getboundingbox is unavailable for this entity type.
  (setq obj (vlax-ename->vla-object ent))
  (if (not (vl-catch-all-error-p
              (vl-catch-all-apply 'vla-getboundingbox
                                  (list obj 'minpt 'maxpt))))
    (progn
      (setq minpt (vlax-safearray->list minpt)
            maxpt (vlax-safearray->list maxpt)
            dx    (- (car  maxpt) (car  minpt))
            dy    (- (cadr maxpt) (cadr minpt))
            pad   (* (max dx dy 0.1) 0.75))
      (command "_.ZOOM" "_W"
               (list (- (car  minpt) pad) (- (cadr minpt) pad))
               (list (+ (car  maxpt) pad) (+ (cadr maxpt) pad)))
    )
    (progn
      (setq ss (ssadd ent (ssadd)))
      (command "_.ZOOM" "_O" ss "")
    )
  )
)


;;; ════════════════════════════════════════════════════════════
;;; Main command
;;; ════════════════════════════════════════════════════════════

(defun c:FINDREPLACE
       (/ searchstr replstr ss slen idx ent etype j
          match-list total choice textval newtext done)

  (vl-load-com)
  (princ "\n=== FINDREPLACE ===")

  ;; ── Step 1: Search string ──────────────────────────────────────────────
  ;; T allows spaces within the entered string.
  (setq searchstr (getstring T "\nEnter search string: "))

  (if (not (= searchstr ""))

    (progn   ; ← skip everything below on empty input

      ;; ── Step 2: Collect matching entities ───────────────────────────────
      ;; "_X" searches the entire drawing database: model space + all layouts.
      ;; The DXF-0 filter string lists every supported type; anything not in
      ;; that list (including Civil3D AECC_* proxies) is silently skipped.
      (setq match-list '())
      (setq ss (ssget "_X"
                  '((0 . "TEXT,MTEXT,DIMENSION,ATTRIB,ATTDEF,MULTILEADER"))))

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
          (setq match-list (reverse match-list))   ; preserve draw order
        )
      )

      (setq total (length match-list))

      (if (= total 0)

        ;; ── No matches ────────────────────────────────────────────────────
        (princ (strcat "\nNo matches found for: \"" searchstr "\"."))

        ;; ── Matches found ─────────────────────────────────────────────────
        (progn

          ;; Apply selection glow to every match — no color properties touched
          (foreach ent match-list (fr:highlight ent))

          (princ (strcat "\nFound " (itoa total)
                         " match(es). Press Enter to continue..."))
          (getstring "")

          ;; ── Step 3: Replacement string ──────────────────────────────────
          (setq replstr (getstring T "\nEnter replacement string: "))

          ;; ── Step 4: One-by-one review loop ──────────────────────────────
          ;; Each Replace is committed immediately.  ESC at any point keeps
          ;; all replacements already made.
          (setq idx  0
                done nil)

          (while (and (< idx total) (not done))

            (setq ent     (nth idx match-list)
                  etype   (cdr (assoc 0 (entget ent)))
                  textval (fr:gettext ent))

            ;; Zoom in; re-apply glow in case the screen refresh cleared it
            (fr:zooment ent)
            (fr:highlight ent)

            (princ (strcat "\n[" (itoa (1+ idx)) "/" (itoa total)
                           "]  Text: \"" textval "\""))

            ;; initget with no bit-1 flag allows a null (Enter) response.
            ;; getkword is wrapped in vl-catch-all-apply so that ESC —
            ;; which raises a cancellation error — can be caught cleanly
            ;; and distinguished from a plain Enter (which returns nil).
            (initget "Replace Keep")
            (setq choice
              (vl-catch-all-apply 'getkword
                (list "\n  [R]eplace / [K]eep / Enter=Replace / ESC to cancel: ")))

            (cond

              ;; ── ESC — caught error ────────────────────────────────────
              ;; Replacements already committed stay committed.
              ((vl-catch-all-error-p choice)
               (fr:dehighlightfrom match-list idx)
               (princ "\nCancelled. Replacements made so far are saved.")
               (setq done T)
              )

              ;; ── Replace — Enter (nil) or "R" keyword ─────────────────
              ((or (null choice) (= choice "Replace"))
               (setq newtext (fr:strreplaceall searchstr replstr textval))
               (fr:settext ent newtext)
               ;; Expand MTEXT box so the longer replacement doesn't wrap.
               ;; fr:fixmtextwidth calls vla-regen, which clears all
               ;; selection glows; re-apply to the remaining entities.
               (if (= etype "MTEXT")
                 (if (fr:fixmtextwidth ent)
                   (progn
                     (setq j (1+ idx))
                     (while (< j total)
                       (fr:highlight (nth j match-list))
                       (setq j (1+ j))
                     )
                   )
                 )
               )
               (fr:dehighlight ent)
               (setq idx (1+ idx))
              )

              ;; ── Keep — "K" keyword ────────────────────────────────────
              ((= choice "Keep")
               (fr:dehighlight ent)
               (setq idx (1+ idx))
              )

            ) ; cond
          ) ; while

          (if (not done)
            (princ "\nFind & Replace complete.")
          )

        ) ; progn — matches found
      ) ; if total = 0

    ) ; progn — main body

    ;; Empty search string
    (princ "\nNo search string entered.")

  ) ; if not empty

  (princ)   ; suppress the nil echo at the Command prompt
)

(princ "\nfindreplace.lsp loaded.  Type FINDREPLACE to run.")
(princ)
;;; ============================================================
;;; End of findreplace.lsp
;;; ============================================================
