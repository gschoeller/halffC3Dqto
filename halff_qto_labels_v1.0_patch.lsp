;; halff_qto_labels_v1.0_patch.lsp
;; Patch: fixes parenthesis in c:QDUPLICATES and c:QFAILSEARCH
;; Load AFTER halff_qto_labels_v1.0.lsp

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
        (princ "\nOK No duplicates found."))))
  (princ))


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
          (princ "\nQFAILSEARCH ended.")))
          )
  (princ))

;; ===========================================
