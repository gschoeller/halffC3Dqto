;; halff_qto_labels_v1.0_patch.lsp
;; v1.3: add QDUMPLABEL diagnostic + v1.1 fixes
;; SELF-CONTAINED: APPLOADing this file is sufficient.
;; It automatically loads halff_qto_labels_v1.0.lsp if not already loaded.

(if (not (and (fboundp 'c:QRUN) (fboundp 'halff:process-dwg-file)))
  (if (findfile "halff_qto_labels_v1.0.lsp")
    (progn
      (princ "\nLoading halff_qto_labels_v1.0.lsp...")
      (load (findfile "halff_qto_labels_v1.0.lsp")))
    (princ "\nWARN: halff_qto_labels_v1.0.lsp not on support path.")))

;; -----------------------------------------------------------------------
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
(defun halff:text-contains? (needle haystack / n h)
  (setq n (strcase (halff:normalize-ws needle))
        h (strcase (halff:normalize-ws haystack)))
  (not (null (vl-string-search n h))))

;; -----------------------------------------------------------------------
(defun halff:dump-pair (code val)
  (princ (strcat "\n  [" (itoa code) "] "))
  (if (= (type val) 'STR)
    (princ val)
    (princ (vl-princ-to-string val))))

;; -----------------------------------------------------------------------
(defun c:QDUMPLABEL
       (/ en ed pair obj pname res xd app sub-en sub-ed)
  (vl-load-com)
  (princ "\nQDUMPLABEL: select a Civil 3D label entity...")
  (setq en (car (entsel "\nSelect label: ")))
  (if (not en)
    (princ "\nNo entity selected.")
    (progn
      (setq ed (entget en))
      (princ (strcat "\n\n=== ENTITY: " (cdr (assoc 0 ed)) " ==="))
      (if (assoc 5 ed)
        (progn
          (princ "\n    Handle: ")
          (princ (cdr (assoc 5 ed)))
          ))
      (princ "\n\n--- DXF group codes (strings) ---")
      (foreach pair ed
        (if (= (type (cdr pair)) 'STR)
          (halff:dump-pair (car pair) (cdr pair))))
      (setq xd (entget en '("*")) app nil)
      (foreach pair xd
        (cond
          ((= (car pair) -3) (setq app T))
          ((and app (= (car pair) 1001))
           (princ (strcat "\n\n--- XDATA: " (cdr pair) " ---"))
           (setq app (cdr pair)))
          (app (halff:dump-pair (car pair) (cdr pair)))
          ))
      (setq obj (vlax-ename->vla-object en))
      (princ "\n\n--- VLA string properties ---")
      (foreach pname '("TextString" "Text" "Contents" "LabelText"
                       "OverrideText" "UserText" "TextOverride"
                       "LabelTextOverride" "DisplayedText"
                       "TextValue" "Name" "Description" "ObjectName")
        (setq res (vl-catch-all-apply
                    'vlax-get-property (list obj pname)))
        (if (and (not (vl-catch-all-error-p res))
                 res (= (type res) 'STR))
          (progn
            (princ (strcat "\n  " pname ": "))
            (princ res))))
      (setq sub-en (entnext en) sub-ed nil)
      (if sub-en (setq sub-ed (entget sub-en)))
      (if (and sub-en sub-ed
               (assoc 330 sub-ed)
               (equal (cdr (assoc 330 sub-ed)) (cdr (assoc 5 ed))))
        (progn
          (princ "\n\n--- First child entity ---")
          (princ (strcat "\n  Type: " (cdr (assoc 0 sub-ed))))
          (foreach pair sub-ed
            (if (= (type (cdr pair)) 'STR)
              (halff:dump-pair (car pair) (cdr pair))))
          ))
      (princ "\n\n--- vlax-dump-object ---")
      (vlax-dump-object obj T)
      (princ "\n\n--- Civil3D method probes ---")
      (foreach meth '("GetTextComponentCount"
                      "GetTextComponentStringAt"
                      "GetOverrideText"
                      "GetUserTextOverride"
                      "GetTextString"
                      "GetLabelTextOverride"
                      "GetLabelDisplayString")
        (setq res (vl-catch-all-apply
                    'vlax-invoke-method (list obj meth 0)))
        (if (not (vl-catch-all-error-p res))
          (progn
            (princ (strcat "\n  " meth "(0): "))
            (princ (vl-princ-to-string res)))
          (progn
            (setq res (vl-catch-all-apply
                        'vlax-invoke-method (list obj meth)))
            (if (not (vl-catch-all-error-p res))
              (progn
                (princ (strcat "\n  " meth "(): "))
                (princ (vl-princ-to-string res))))
            )))
      (princ "\n\n=== QDUMPLABEL done ===\n")))
  (princ))

;; -----------------------------------------------------------------------
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
