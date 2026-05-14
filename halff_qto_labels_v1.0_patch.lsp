;; halff_qto_labels_v1.0_patch.lsp
;; v1.1 fixes:
;;   1. Whitespace-normalized text match (multi-line MLEADER/MTEXT)
;;   2. DXF group-code 1/3 fallback for Civil 3D label text extraction
;;   3. Correct parenthesis in c:QDUPLICATES and c:QFAILSEARCH
;; Load AFTER halff_qto_labels_v1.0.lsp

;; -----------------------------------------------------------------------
;; Normalize whitespace: collapses space/tab/LF/CR and MTEXT \P paragraph
;; breaks (backslash=92 + P=80 or p=112) into a single space.
(defun halff:normalize-ws (s / i c code out last-sp)
  (if (or (null s) (= s "")) ""
    (progn
      (setq out "" i 1 last-sp T)
      (while (<= i (strlen s))
        (setq c (substr s i 1)
              code (ascii c))
        (cond
          ((and (= code 92)
                (<= (1+ i) (strlen s))
                (member (ascii (substr s (1+ i) 1)) '(80 112)))
           (if (not last-sp)
             (setq out (strcat out " ") last-sp T))
           (setq i (1+ i)))
          ((member code '(9 10 13 32))
           (if (not last-sp)
             (setq out (strcat out " ") last-sp T)))
          (T (setq out (strcat out c) last-sp nil)))
        (setq i (1+ i)))
      (setq out (vl-string-trim " " out))
      out)))

;; -----------------------------------------------------------------------
;; Extract text via DXF group codes 1 and 3 (fallback for Civil 3D labels).
(defun halff:dxf-get-text (en / ed texts piece result)
  (setq ed (entget en) texts '())
  (foreach pair ed
    (if (member (car pair) '(1 3))
      (setq texts (cons (cdr pair) texts))))
  (if texts
    (progn
      (setq result "")
      (foreach piece (reverse texts)
        (setq result (strcat result piece)))
      (if (= result "") nil result))
    nil))

;; -----------------------------------------------------------------------
;; Extract display text from any entity.
;; Tries vla-get-textstring, then alternate VLA properties, then DXF codes.
(defun halff:vla-get-text (obj / res en props txt)
  (setq res (vl-catch-all-apply 'vla-get-textstring (list obj)))
  (if (and (not (vl-catch-all-error-p res)) res (not (= res "")))
    res
    (progn
      (setq props '("TextString" "Text" "Contents" "LabelText")
            txt   nil)
      (while (and props (not txt))
        (setq res (vl-catch-all-apply
                    'vlax-get-property (list obj (car props))))
        (if (and (not (vl-catch-all-error-p res))
                 res (= (type res) 'STR) (not (= res "")))
          (setq txt res))
        (setq props (cdr props)))
      (if (not txt)
        (progn
          (setq en (vl-catch-all-apply
                     'vlax-vla-object->ename (list obj)))
          (if (not (vl-catch-all-error-p en))
            (setq txt (halff:dxf-get-text en)))
          ))
      txt)))

;; -----------------------------------------------------------------------
;; Case-insensitive substring match with whitespace normalization.
(defun halff:text-contains? (needle haystack / n h)
  (setq n (strcase (halff:normalize-ws needle))
        h (strcase (halff:normalize-ws haystack)))
  (not (null (vl-string-search n h))))

;; -----------------------------------------------------------------------
;; QDUPLICATES - correct parenthesis version
(defun c:QDUPLICATES (/ tol prec s i e key seen pair dup cnt)
  (vl-load-com)
  (initget 6)
  (setq tol (getreal "\nTolerance for duplicate detection <0.001>: "))
  (if (not tol) (setq tol 0.001))
  (setq prec (halff:tol->prec tol))
  (setq s (ssget "_X"))
  (if (not s)
    (princ "\nX No objects found.")
    (progn
      (setq seen '())
      (setq dup (ssadd))
      (setq cnt 0)
      (setq i 0)
      (while (< i (sslength s))
        (setq e (ssname s i))
        (setq key (halff:dup-key e prec))
        (setq pair (assoc key seen))
        (if pair
          (progn
            (if (not (ssmemb (cdr pair) dup)) (ssadd (cdr pair) dup))
            (if (not (ssmemb e dup)) (ssadd e dup))
            (setq cnt (1+ cnt)))
          (setq seen (cons (cons key e) seen)))
        (setq i (1+ i)))
      (if (> (sslength dup) 0)
        (progn
          (sssetfirst nil dup)
          (princ (strcat "\nOK Found " (itoa (sslength dup))
                         " duplicated objects. Selected duplicates.")))
        (princ "\nOK No duplicates found."))
      ))
  (princ))

;; -----------------------------------------------------------------------
;; QFAILSEARCH - correct parenthesis version
(defun c:QFAILSEARCH (/ ens idx n en obj prevObj ss cmd)
  (vl-load-com)
  (if (not *HALFF_FAIL_ENTS*)
    (princ "\nX No failure data found. Run QRUN first.")
    (progn
      (setq ens *HALFF_FAIL_ENTS*)
      (setq n (length ens))
      (if (= n 0)
        (princ "\nX No failure entities cached for current drawing.")
        (progn
          (princ (strcat "\n[QFAILSEARCH] " (itoa n)
                         " failure entities. Use Next/Previous/Quit."))
          (setq idx 0)
          (setq prevObj nil)
          (while (and (>= idx 0) (< idx n))
            (setq en (nth idx ens))
            (if prevObj
              (vl-catch-all-apply 'vla-Highlight (list prevObj :vlax-false)))
            (if (and en (entget en))
              (progn
                (setq obj (vlax-ename->vla-object en))
                (halff:zoom-to-ename en)
                (vl-catch-all-apply 'vla-Highlight (list obj :vlax-true))
                (setq prevObj obj)
                (sssetfirst nil nil)
                (setq ss (ssadd))
                (ssadd en ss)
                (sssetfirst nil ss)
                (princ (strcat "\n[" (itoa (1+ idx)) "/" (itoa n)
                               "] FAILURE ENTITY"
                               "  (N)ext / (P)revious / (Q)uit"))
                (initget "Next Previous Quit")
                (setq cmd (getkword "\nNext/Previous/Quit <Next>: "))
                (cond
                  ((or (not cmd) (= cmd "Next"))
                   (setq idx (if (< idx (1- n)) (1+ idx) 0)))
                  ((= cmd "Previous")
                   (setq idx (if (> idx 0) (1- idx) (1- n))))
                  ((= cmd "Quit")
                   (progn
                     (if prevObj
                       (vl-catch-all-apply 'vla-Highlight
                                           (list prevObj :vlax-false)))
                     (sssetfirst nil nil)
                     (setq idx -1))))
                     )
              (setq idx (if (< idx (1- n)) (1+ idx) 0))))
          (if prevObj
            (vl-catch-all-apply 'vla-Highlight (list prevObj :vlax-false)))
          (princ "\nQFAILSEARCH ended."))
        ))
      )
  (princ))

;; ===========================================
