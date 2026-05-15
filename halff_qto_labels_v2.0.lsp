;; =====================================================================
;; Halff QTO Labels v2.0
;;
;; Built on Halff QTO Labels v1.0.  Adds TEXT QTY column support so
;; that QRUN can extract and sum numeric quantities written as AutoCAD
;; TEXT / MTEXT entities rather than as block attributes.
;;
;; Compatibility: AutoCAD 2018+ (tested on 2022 and 2024)
;; Dependencies : none (stand-alone LSP)
;;
;; Public commands
;;   QLABEL   - interactive label-placement wizard
;;   QRUN     - batch-process all labelled drawings
;;   QREPORT  - write summary CSV
;;   QCHECK   - audit labels in current drawing
;;   QFAILSEARCH - step through QRUN failure entities
;; =====================================================================

(vl-load-com)

;;;; ----------------------------------------------------------------
;;;; 0.  Global constants & version stamp
;;;; ----------------------------------------------------------------

(setq *HQTO-VERSION* "2.0"
      *HQTO-LABEL-LAYER* "HQTO_LABELS"
      *HQTO-RESULT-LAYER* "HQTO_RESULTS"
      *HQTO-LOG-FILE*    (strcat (getvar "TEMPPREFIX") "halff_qto_run.log")
      *HQTO-CSV-FILE*    (strcat (getvar "TEMPPREFIX") "halff_qto_report.csv")
      *HQTO-DWG-LIST*    nil      ; populated by QRUN
      *HQTO-FAIL-LIST*   nil      ; entities that could not be processed
      *HQTO-FAIL-IDX*    0
)

;;;; ----------------------------------------------------------------
;;;; 1.  Utility helpers
;;;; ----------------------------------------------------------------

(defun hqto:log (msg / fp)
  "Append MSG to the session log file."
  (setq fp (open *HQTO-LOG-FILE* "a"))
  (write-line (strcat (rtos (getvar "DATE") 2 8) "  " msg) fp)
  (close fp)
)

(defun hqto:trim (s)
  "Strip leading/trailing whitespace from string S."
  (vl-string-trim '(32 9 10 13) s)
)

(defun hqto:str->real (s / trimmed)
  "Convert string S to real; return nil if not numeric."
  (setq trimmed (hqto:trim s))
  (if (wcmatch trimmed "*[0-9]*")
    (distof trimmed 2)
    nil
  )
)

(defun hqto:layer-exists-p (lname)
  (tblsearch "LAYER" lname)
)

(defun hqto:make-layer (lname color / cmd-echo)
  (setq cmd-echo (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (command "_.LAYER" "_Make" lname "_Color" color lname "")
  (setvar "CMDECHO" cmd-echo)
)

(defun hqto:ensure-layers ()
  (unless (hqto:layer-exists-p *HQTO-LABEL-LAYER*)
    (hqto:make-layer *HQTO-LABEL-LAYER* 3)   ; green
  )
  (unless (hqto:layer-exists-p *HQTO-RESULT-LAYER*)
    (hqto:make-layer *HQTO-RESULT-LAYER* 1)  ; red
  )
)

(defun hqto:xdata-key (ename)
  "Return the HQTO xdata list attached to ENAME, or nil."
  (cdr (assoc -3 (entget ename '("HQTO"))))
)

(defun hqto:attach-xdata (ename key value / regapp-ok xd)
  "Attach/replace HQTO xdata on ENAME with KEY=VALUE (both strings)."
  (setq regapp-ok (regapp "HQTO"))
  (setq xd
    (list
      (cons -3
        (list
          (list "HQTO"
            (cons 1000 key)
            (cons 1000 value)
          )
        )
      )
    )
  )
  (entmod (append (entget ename) xd))
)

(defun hqto:get-xdata-value (ename / xd pairs)
  "Return the value string stored in HQTO xdata on ENAME."
  (setq xd (hqto:xdata-key ename))
  (if xd
    (progn
      (setq pairs (cdar xd))
      (cdr (assoc 1000 (cdr pairs)))
    )
    nil
  )
)

;;;; ----------------------------------------------------------------
;;;; 2.  Geometry helpers
;;;; ----------------------------------------------------------------

(defun hqto:ent-center (ename / ed etype)
  "Return insertion / center point of entity ENAME as 2-D point."
  (setq ed    (entget ename)
        etype (cdr (assoc 0 ed))
  )
  (cond
    ((member etype '("INSERT" "TEXT" "MTEXT"))
     (cdr (assoc 10 ed))
    )
    ((= etype "CIRCLE")
     (cdr (assoc 10 ed))
    )
    ((= etype "LINE")
     (mapcar '(lambda (a b) (/ (+ a b) 2.0))
             (cdr (assoc 10 ed))
             (cdr (assoc 11 ed))
     )
    )
    (t
     ;; fallback: bounding-box centre via GETBOUNDINGBOX
     (vla-getboundingbox
       (vlax-ename->vla-object ename)
       'mn 'mx
     )
     (mapcar '(lambda (a b) (/ (+ a b) 2.0))
             (vlax-safearray->list mn)
             (vlax-safearray->list mx)
     )
    )
  )
)

(defun hqto:pt-in-poly-p (pt poly-pts / n i j c px py xi xj yi yj)
  "Ray-casting point-in-polygon test.  poly-pts is a flat list of
   2-D points forming a closed polygon (last pt != first pt ok)."
  (setq n (length poly-pts)
        i 0  j (1- n)  c nil
        px (car pt)  py (cadr pt)
  )
  (while (< i n)
    (setq xi (car  (nth i poly-pts))
          yi (cadr (nth i poly-pts))
          xj (car  (nth j poly-pts))
          yj (cadr (nth j poly-pts))
    )
    (if (and
          (/= (> yi py) (> yj py))
          (< px (+ xi (* (/ (- xj xi) (- yj yi)) (- py yi))))
        )
      (setq c (not c))
    )
    (setq j i  i (1+ i))
  )
  c
)

;;;; ----------------------------------------------------------------
;;;; 3.  Viewport / Sheet helpers (for paper-space label matching)
;;;; ----------------------------------------------------------------

(defun hqto:all-viewports (/ ss i vp-list)
  "Return a list of all VIEWPORT enames in the current layout."
  (setq vp-list nil)
  (if (setq ss (ssget "X" '((0 . "VIEWPORT"))))
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq vp-list (cons (ssname ss i) vp-list)
              i       (1+ i)
        )
      )
    )
  )
  vp-list
)

(defun hqto:vp-poly (vpEnt / ed cx cy hw hh ang cos-a sin-a corners)
  "Return a 4-point polygon (list of 2-D pts) representing the
   boundary of viewport VPENT in paper space."
  (setq ed  (entget vpEnt)
        cx  (car  (cdr (assoc 10 ed)))
        cy  (cadr (cdr (assoc 10 ed)))
        hw  (/ (cdr (assoc 40 ed)) 2.0)   ; half-width
        hh  (/ (cdr (assoc 41 ed)) 2.0)   ; half-height
        ang (if (assoc 51 ed) (cdr (assoc 51 ed)) 0.0)
        cos-a (cos ang)
        sin-a (sin ang)
  )
  (defun rot (dx dy)
    (list (+ cx (- (* dx cos-a) (* dy sin-a)))
          (+ cy (+ (* dx sin-a) (* dy cos-a)))
    )
  )
  (list (rot (- hw) (- hh))
        (rot    hw  (- hh))
        (rot    hw     hh)
        (rot (- hw)    hh)
  )
)

(defun hqto:label-in-vp (labelPt vpEnt)
  "Return T if 2-D point LABELPT falls inside the boundary of VPENT."
  (hqto:pt-in-poly-p labelPt (hqto:vp-poly vpEnt))
)

(defun hqto:find-vp-for-label (labelPt / vps result)
  "Return the VIEWPORT ename that contains LABELPT, or nil."
  (setq vps    (hqto:all-viewports)
        result nil
  )
  (while (and vps (not result))
    (if (hqto:label-in-vp labelPt (car vps))
      (setq result (car vps))
    )
    (setq vps (cdr vps))
  )
  result
)

;;;; ----------------------------------------------------------------
;;;; 4.  Text-quantity extraction  (NEW in v2.0)
;;;; ----------------------------------------------------------------

(defun hqto:mtext-plain (ename / raw)
  "Strip basic MTEXT formatting codes and return plain string."
  (setq raw (cdr (assoc 1 (entget ename))))
  ;; Remove \P (paragraph), \n, \t, {\...} formatting groups
  (setq raw (vl-string-subst "" "\\P" raw))
  (setq raw (vl-string-subst " " "\\n" raw))
  (setq raw (vl-string-subst " " "\\t" raw))
  ;; Strip { } formatting blocks: naive single-pass removal
  (while (vl-string-search "{" raw)
    (setq raw
      (vl-regex-subst "{[^}]*}" "" raw)
    )
  )
  (hqto:trim raw)
)

(defun hqto:entity-text (ename / etype ed)
  "Return the display string of TEXT or MTEXT entity ENAME."
  (setq etype (cdr (assoc 0 (entget ename))))
  (cond
    ((= etype "TEXT")  (cdr (assoc 1 (entget ename))))
    ((= etype "MTEXT") (hqto:mtext-plain ename))
    (t nil)
  )
)

(defun hqto:collect-text-qty (regionPts / ss i ename etype txt val total)
  "Collect all TEXT/MTEXT entities whose insertion point lies inside
   REGIONPTS (2-D polygon) and sum their numeric content.
   Returns (total . list-of-enames)."
  (setq ss    (ssget "CP" regionPts '((0 . "TEXT,MTEXT")))
        total 0.0
        result-ents nil
  )
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ename (ssname ss i)
              txt   (hqto:entity-text ename)
              val   (if txt (hqto:str->real txt) nil)
        )
        (if val
          (progn
            (setq total       (+ total val)
                  result-ents (cons ename result-ents)
            )
          )
          (progn
            (hqto:log (strcat "  SKIP non-numeric text: " (if txt txt "<nil>")))
          )
        )
        (setq i (1+ i))
      )
    )
  )
  (cons total result-ents)
)

;;;; ----------------------------------------------------------------
;;;; 5.  Label-block attribute helpers
;;;; ----------------------------------------------------------------

(defun hqto:get-attr (blockEnt tag / atts a)
  "Return value of attribute TAG on INSERT entity BLOCKENT."
  (setq atts (hqto:block-atts blockEnt))
  (setq a (assoc (strcase tag) atts))
  (if a (cdr a) nil)
)

(defun hqto:set-attr (blockEnt tag value / obj atts a)
  "Set attribute TAG on INSERT entity BLOCKENT to VALUE."
  (setq obj  (vlax-ename->vla-object blockEnt)
        atts (vlax-invoke obj 'GetAttributes)
  )
  (foreach att atts
    (if (= (strcase (vlax-get att 'TagString)) (strcase tag))
      (vlax-put att 'TextString value)
    )
  )
)

(defun hqto:block-atts (blockEnt / obj atts result)
  "Return alist of (TAG . VALUE) for all attributes of INSERT BLOCKENT."
  (setq obj    (vlax-ename->vla-object blockEnt)
        atts   (vlax-invoke obj 'GetAttributes)
        result nil
  )
  (foreach a atts
    (setq result
      (cons (cons (strcase (vlax-get a 'TagString))
                  (vlax-get a 'TextString)
            )
            result
      )
    )
  )
  result
)

;;;; ----------------------------------------------------------------
;;;; 6.  Label record structure
;;;;
;;;;   A "label" is a block INSERT on layer HQTO_LABELS that carries
;;;;   these attributes (tag names are case-insensitive):
;;;;     QTY_TYPE : "ATTR" | "TEXT"   (default "ATTR" for v1 compat.)
;;;;     CATEGORY : string
;;;;     ITEM     : string
;;;;     QTY      : numeric string (written by QRUN)
;;;;     UNIT     : string
;;;;     NOTE     : optional
;;;; ----------------------------------------------------------------

(defun hqto:label-qty-type (labelEnt)
  (or (hqto:get-attr labelEnt "QTY_TYPE") "ATTR")
)

(defun hqto:label-region-pts (labelEnt / ed ins w h ang cos-a sin-a)
  "Derive a rectangular region from the label block's bounding box
   scaled by the REGION_SCALE attribute (default 1.0).
   Returns a list of 4 2-D points."
  (setq ed    (entget labelEnt)
        ins   (cdr (assoc 10 ed))
        w     (hqto:str->real (or (hqto:get-attr labelEnt "REGION_W") "1000"))
        h     (hqto:str->real (or (hqto:get-attr labelEnt "REGION_H") "1000"))
        ang   (if (assoc 50 ed) (cdr (assoc 50 ed)) 0.0)
        cos-a (cos ang)
        sin-a (sin ang)
  )
  ;; Build rotated rectangle centred on insertion point
  (defun rpt (dx dy)
    (list (+ (car ins) (- (* dx cos-a) (* dy sin-a)))
          (+ (cadr ins) (+ (* dx sin-a) (* dy cos-a)))
    )
  )
  (list (rpt (/ w -2.0) (/ h -2.0))
        (rpt (/ w  2.0) (/ h -2.0))
        (rpt (/ w  2.0) (/ h  2.0))
        (rpt (/ w -2.0) (/ h  2.0))
  )
)

;;;; ----------------------------------------------------------------
;;;; 7.  Quantity computation dispatcher
;;;; ----------------------------------------------------------------

(defun hqto:compute-qty (labelEnt / qtype result)
  "Compute and return the quantity for LABELENT.
   For QTY_TYPE=ATTR: count matching block attributes in region.
   For QTY_TYPE=TEXT: sum numeric TEXT/MTEXT strings in region.
   Returns a real number."
  (setq qtype (strcase (hqto:label-qty-type labelEnt)))
  (cond
    ((= qtype "TEXT")
     (hqto:compute-qty-text labelEnt)
    )
    (t  ; default ATTR
     (hqto:compute-qty-attr labelEnt)
    )
  )
)

(defun hqto:compute-qty-text (labelEnt / region res)
  "Sum TEXT/MTEXT numeric values inside the label's region."
  (setq region (hqto:label-region-pts labelEnt)
        res    (hqto:collect-text-qty region)
  )
  (hqto:log (strcat "  TEXT QTY total=" (rtos (car res) 2 4)
                    " ents=" (itoa (length (cdr res)))
            )
  )
  (car res)
)

(defun hqto:compute-qty-attr (labelEnt / region category item ss i cnt)
  "Count block insertions matching CATEGORY/ITEM inside the label's region."
  (setq region   (hqto:label-region-pts labelEnt)
        category (strcase (or (hqto:get-attr labelEnt "CATEGORY") ""))
        item     (strcase (or (hqto:get-attr labelEnt "ITEM") ""))
        ss       (ssget "CP" region '((0 . "INSERT")))
        cnt      0
  )
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ename (ssname ss i))
        (if (hqto:block-matches-p ename category item)
          (setq cnt (1+ cnt))
        )
        (setq i (1+ i))
      )
    )
  )
  (float cnt)
)

(defun hqto:block-matches-p (ename category item / atts cat itm)
  "Return T if INSERT ENAME has CATEGORY=CATEGORY and ITEM=ITEM."
  (setq atts (hqto:block-atts ename)
        cat  (strcase (or (cdr (assoc "CATEGORY" atts)) ""))
        itm  (strcase (or (cdr (assoc "ITEM"     atts)) ""))
  )
  (and (= cat category) (= itm item))
)

;;;; ----------------------------------------------------------------
;;;; 8.  QLABEL command  -- interactive label-placement wizard
;;;; ----------------------------------------------------------------

(defun C:QLABEL ( / qtype category item unit note region-w region-h
                    ins-pt blk-name ent )
  (hqto:ensure-layers)
  (initget "ATTRibute TEXT")
  (setq qtype (getkword "\nQuantity type [ATTRibute/TEXT] <ATTRibute>: "))
  (if (null qtype) (setq qtype "ATTR"))
  (setq qtype (if (= (strcase qtype 1) "text") "TEXT" "ATTR"))

  (setq category (getstring T "\nCategory <GENERAL>: "))
  (if (= category "") (setq category "GENERAL"))

  (setq item (getstring T "\nItem name: "))
  (while (= item "")
    (setq item (getstring T "  Item cannot be blank.  Item name: "))
  )

  (setq unit (getstring T "\nUnit (e.g. EA, LF, SF) <EA>: "))
  (if (= unit "") (setq unit "EA"))

  (setq region-w
    (getdist (strcat "\nRegion width in drawing units <1000>: "))
  )
  (if (null region-w) (setq region-w 1000.0))

  (setq region-h
    (getdist (strcat "\nRegion height in drawing units <1000>: "))
  )
  (if (null region-h) (setq region-h 1000.0))

  (setq note (getstring T "\nOptional note <blank>: "))

  (setq ins-pt (getpoint "\nPick label insertion point: "))
  (while (null ins-pt)
    (setq ins-pt (getpoint "  Pick a point: "))
  )

  ;; Insert the label block (must exist in drawing or search path)
  (setq blk-name "HQTO_LABEL")
  (setvar "CLAYER" *HQTO-LABEL-LAYER*)
  (command "_.INSERT" blk-name ins-pt 1 1 0)

  ;; Locate the newly inserted block
  (setq ent (entlast))

  ;; Populate attributes
  (hqto:set-attr ent "QTY_TYPE"  qtype)
  (hqto:set-attr ent "CATEGORY"  category)
  (hqto:set-attr ent "ITEM"      item)
  (hqto:set-attr ent "UNIT"      unit)
  (hqto:set-attr ent "REGION_W"  (rtos region-w 2 4))
  (hqto:set-attr ent "REGION_H"  (rtos region-h 2 4))
  (hqto:set-attr ent "NOTE"      note)
  (hqto:set-attr ent "QTY"       "0")

  (hqto:attach-xdata ent "LABEL_ID"
    (strcat category "-" item "-"
            (itoa (fix (getvar "DATE")))
    )
  )

  (princ (strcat "\nLabel placed for " category "/" item
                 " (" qtype " mode)."
         )
  )
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 9.  QCHECK command  -- audit labels in current drawing
;;;; ----------------------------------------------------------------

(defun C:QCHECK ( / ss i ename atts ok-cnt err-cnt)
  (setq ss      (ssget "X" (list (cons 0 "INSERT")
                                 (cons 8 *HQTO-LABEL-LAYER*)
                          )
               )
        ok-cnt  0
        err-cnt 0
  )
  (if (null ss)
    (progn
      (princ "\nNo HQTO labels found in current drawing.")
      (princ)
      (exit)
    )
  )
  (princ (strcat "\nChecking " (itoa (sslength ss)) " labels..."))
  (setq i 0)
  (while (< i (sslength ss))
    (setq ename (ssname ss i)
          atts  (hqto:block-atts ename)
    )
    (if (and (assoc "CATEGORY" atts)
             (assoc "ITEM"     atts)
             (assoc "UNIT"     atts)
        )
      (setq ok-cnt (1+ ok-cnt))
      (progn
        (setq err-cnt (1+ err-cnt))
        (hqto:log (strcat "QCHECK ERR ename=" (vl-prin1-to-string ename)))
        (princ (strcat "\n  [ERR] Label at "
                       (vl-prin1-to-string (cdr (assoc 10 (entget ename))))
                       " missing required attributes."
               )
        )
      )
    )
    (setq i (1+ i))
  )
  (princ (strcat "\nQCHECK complete: " (itoa ok-cnt) " OK, "
                 (itoa err-cnt) " errors."
         )
  )
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 10. QRUN command  -- batch-process drawings
;;;; ----------------------------------------------------------------

(defun C:QRUN ( / dwg-dir file-list)
  (setq *HQTO-FAIL-LIST* nil
        *HQTO-FAIL-IDX*  0
  )
  (hqto:log (strcat "===== QRUN start v" *HQTO-VERSION* " ====="))

  ;; Ask user for folder
  (setq dwg-dir
    (getfiled "Select any DWG in the target folder" "" "dwg" 2)
  )
  (if (null dwg-dir)
    (progn (princ "\nQRUN cancelled.") (princ) (exit))
  )
  (setq dwg-dir (vl-filename-directory dwg-dir))

  ;; Enumerate DWGs
  (setq file-list
    (vl-directory-files dwg-dir "*.dwg" 1)
  )
  (if (null file-list)
    (progn (princ "\nNo DWG files found.") (princ) (exit))
  )

  (setq *HQTO-DWG-LIST* file-list)
  (princ (strcat "\nProcessing " (itoa (length file-list))
                 " drawings in " dwg-dir "..."
         )
  )

  (foreach dwg file-list
    (hqto:process-drawing (strcat dwg-dir "\\" dwg))
  )

  (princ (strcat "\nQRUN complete.  "
                 (itoa (length *HQTO-FAIL-LIST*))
                 " failures logged."
         )
  )
  (hqto:log "===== QRUN end =====")
  (princ)
)

(defun hqto:process-drawing (dwg-path / old-dwg ss i labelEnt qty err)
  "Open DWG-PATH read/write, update all label QTY attributes, save."
  (hqto:log (strcat "Processing: " dwg-path))
  (setq old-dwg (getvar "DWGNAME"))
  ;; Use OPEN approach via vl-file-copy safety not needed; direct open
  (setq err
    (vl-catch-all-apply
      '(lambda ()
          (command "_.OPEN" dwg-path)
          (hqto:update-labels-in-current-dwg dwg-path)
          (command "_.QSAVE")
       )
    )
  )
  (if (vl-catch-all-error-p err)
    (progn
      (hqto:log (strcat "  ERROR: " (vl-catch-all-error-message err)))
      (setq *HQTO-FAIL-LIST*
        (cons (list dwg-path (vl-catch-all-error-message err))
              *HQTO-FAIL-LIST*
        )
      )
    )
  )
)

(defun hqto:update-labels-in-current-dwg (dwg-path / ss i ename qty)
  "Update QTY attribute on every HQTO label in the current drawing."
  (setq ss (ssget "X" (list (cons 0 "INSERT")
                            (cons 8 *HQTO-LABEL-LAYER*)
                    )
           )
  )
  (if (null ss)
    (progn
      (hqto:log "  No labels found.")
      (exit)
    )
  )
  (princ (strcat "\n    Processing CURRENT drawing"))
  (setq i 0)
  (while (< i (sslength ss))
    (setq ename (ssname ss i))
    (setq qty
      (vl-catch-all-apply 'hqto:compute-qty (list ename))
    )
    (if (vl-catch-all-error-p qty)
      (progn
        (hqto:log
          (strcat "  FAIL label @ "
                  (vl-prin1-to-string (cdr (assoc 10 (entget ename))))
                  " : " (vl-catch-all-error-message qty)
          )
        )
        (setq *HQTO-FAIL-LIST*
          (cons (list dwg-path ename (vl-catch-all-error-message qty))
                *HQTO-FAIL-LIST*
          )
        )
      )
      (progn
        (hqto:set-attr ename "QTY" (rtos qty 2 4))
        (hqto:log
          (strcat "  OK  qty=" (rtos qty 2 4))
        )
      )
    )
    (setq i (1+ i))
  )
)

;;;; ----------------------------------------------------------------
;;;; 11. QREPORT command  -- write CSV summary
;;;; ----------------------------------------------------------------

(defun C:QREPORT ( / fp dwg-dir file-list out-path)
  (setq out-path
    (getfiled "Save QTO report as" *HQTO-CSV-FILE* "csv" 1)
  )
  (if (null out-path)
    (progn (princ "\nQREPORT cancelled.") (princ) (exit))
  )
  (setq fp (open out-path "w"))
  (write-line "Drawing,Layer,Category,Item,QTY,Unit,Note" fp)

  (setq dwg-dir
    (getfiled "Select any DWG in the target folder" "" "dwg" 2)
  )
  (if (null dwg-dir)
    (progn (close fp) (princ "\nCancelled.") (princ) (exit))
  )
  (setq dwg-dir   (vl-filename-directory dwg-dir)
        file-list (vl-directory-files dwg-dir "*.dwg" 1)
  )
  (foreach dwg file-list
    (hqto:report-one-drawing (strcat dwg-dir "\\" dwg) fp)
  )
  (close fp)
  (princ (strcat "\nReport written to " out-path))
  (princ)
)

(defun hqto:report-one-drawing (dwg-path fp / ss i ename atts row)
  (command "_.OPEN" dwg-path)
  (setq ss (ssget "X" (list (cons 0 "INSERT")
                            (cons 8 *HQTO-LABEL-LAYER*)
                    )
           )
  )
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ename (ssname ss i)
              atts  (hqto:block-atts ename)
        )
        (setq row
          (strcat
            (vl-filename-base dwg-path) ","
            *HQTO-LABEL-LAYER* ","
            (or (cdr (assoc "CATEGORY" atts)) "") ","
            (or (cdr (assoc "ITEM"     atts)) "") ","
            (or (cdr (assoc "QTY"      atts)) "0") ","
            (or (cdr (assoc "UNIT"     atts)) "") ","
            (or (cdr (assoc "NOTE"     atts)) "")
          )
        )
        (write-line row fp)
        (setq i (1+ i))
      )
    )
  )
)

;;;; ----------------------------------------------------------------
;;;; 12. QFAILSEARCH command  -- step through failure entities
;;;; ----------------------------------------------------------------

(defun C:QFAILSEARCH ( / entry dwg-path ename msg)
  (if (null *HQTO-FAIL-LIST*)
    (progn
      (princ "\nNo failures recorded in this session.")
      (princ)
      (exit)
    )
  )
  (if (>= *HQTO-FAIL-IDX* (length *HQTO-FAIL-LIST*))
    (setq *HQTO-FAIL-IDX* 0)
  )
  (setq entry    (nth *HQTO-FAIL-IDX* *HQTO-FAIL-LIST*)
        dwg-path (car entry)
        ename    (cadr entry)
        msg      (caddr entry)
  )
  (princ (strcat "\nFailure " (itoa (1+ *HQTO-FAIL-IDX*))
                 " of " (itoa (length *HQTO-FAIL-LIST*)) ":"
         )
  )
  (princ (strcat "\n  Drawing : " dwg-path))
  (if (= (type ename) 'ENAME)
    (progn
      (princ (strcat "\n  Entity  : " (vl-prin1-to-string ename)))
      ;; Zoom to entity if in same drawing
      (if (= (getvar "DWGNAME")
             (vl-filename-base dwg-path)
          )
        (command "_.ZOOM" "_Object" ename "")
      )
    )
    (princ (strcat "\n  Detail  : " (if msg msg "(no detail)")))
  )
  (setq *HQTO-FAIL-IDX* (1+ *HQTO-FAIL-IDX*))
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 13. Extended-data registration on load
;;;; ----------------------------------------------------------------

(regapp "HQTO")
(hqto:log (strcat "halff_qto_labels v" *HQTO-VERSION* " loaded."))

;;;; ----------------------------------------------------------------
;;;; 14. Multi-drawing viewport-aware processing (paper-space labels)
;;;; ----------------------------------------------------------------

(defun hqto:update-ps-labels-in-current-dwg (dwg-path / ss i ename vpEnt hitCount vps matchVP qty)
  "Variant of hqto:update-labels-in-current-dwg for drawings that
   use paper-space labels.  Each label is matched to the viewport
   that contains its insertion point; the quantity is computed using
   the model-space region corresponding to that viewport."
  (setq ss (ssget "X" (list (cons 0 "INSERT")
                            (cons 8 *HQTO-LABEL-LAYER*)
                            (cons 67 1)   ; paper space flag
                    )
           )
  )
  (if (null ss)
    (progn
      (hqto:log "  No paper-space labels found.")
      (exit)
    )
  )
  (setq i 0)
  (while (< i (sslength ss))
    (setq ename    (ssname ss i)
          hitCount 0
          matchVP  nil
          vps      (hqto:all-viewports)
    )
    (while vps
      (if (hqto:label-in-vp (hqto:ent-center ename) (car vps))
        (progn
          (setq hitCount (1+ hitCount)
                matchVP  (car vps)
          )
        )
      )
      (setq vps (cdr vps))
    )
    (cond
      ((= hitCount 1)
       ;; Assign to the one matching VP
       (setq vpEnt matchVP)
      )
      ((> hitCount 1)
       (hqto:log "  WARN: label inside multiple VPs; using first match.")
       (setq vpEnt matchVP)
      )
      (t
       (hqto:log "  WARN: label not inside any VP; skipping.")
       (setq i (1+ i))
       ;; continue to next label
      )
    )
    (setq qty
      (vl-catch-all-apply 'hqto:compute-qty (list ename))
    )
    (if (vl-catch-all-error-p qty)
      (progn
        (hqto:log
          (strcat "  FAIL ps-label @ "
                  (vl-prin1-to-string (cdr (assoc 10 (entget ename))))
                  " : " (vl-catch-all-error-message qty)
          )
        )
        (setq *HQTO-FAIL-LIST*
          (cons (list dwg-path ename (vl-catch-all-error-message qty))
                *HQTO-FAIL-LIST*
          )
        )
      )
      (hqto:set-attr ename "QTY" (rtos qty 2 4))
    )
    (setq i (1+ i))
  )
)

;;;; ----------------------------------------------------------------
;;;; 15. Region-scale attribute support
;;;; ----------------------------------------------------------------

(defun hqto:label-region-pts-v2 (labelEnt / ed ins scl w h ang cos-a sin-a)
  "Extended version: honours REGION_SCALE multiplier attribute."
  (setq ed    (entget labelEnt)
        ins   (cdr (assoc 10 ed))
        scl   (hqto:str->real (or (hqto:get-attr labelEnt "REGION_SCALE") "1.0"))
        w     (* (hqto:str->real (or (hqto:get-attr labelEnt "REGION_W") "1000"))
                 (if scl scl 1.0)
              )
        h     (* (hqto:str->real (or (hqto:get-attr labelEnt "REGION_H") "1000"))
                 (if scl scl 1.0)
              )
        ang   (if (assoc 50 ed) (cdr (assoc 50 ed)) 0.0)
        cos-a (cos ang)
        sin-a (sin ang)
  )
  (defun rpt (dx dy)
    (list (+ (car ins) (- (* dx cos-a) (* dy sin-a)))
          (+ (cadr ins) (+ (* dx sin-a) (* dy cos-a)))
    )
  )
  (list (rpt (/ w -2.0) (/ h -2.0))
        (rpt (/ w  2.0) (/ h -2.0))
        (rpt (/ w  2.0) (/ h  2.0))
        (rpt (/ w -2.0) (/ h  2.0))
  )
)

;;;; ----------------------------------------------------------------
;;;; 16. Unit-conversion helper
;;;; ----------------------------------------------------------------

(defun hqto:convert-unit (value from-unit to-unit / factor)
  "Convert VALUE from FROM-UNIT to TO-UNIT.  Returns converted real.
   Supported units: IN FT YD MM CM M SF SY SF-to-SY LF-to-M etc."
  (setq factor
    (cond
      ;; Length
      ((and (= from-unit "IN")  (= to-unit "FT"))  (/ 1.0 12.0))
      ((and (= from-unit "FT")  (= to-unit "IN"))  12.0)
      ((and (= from-unit "FT")  (= to-unit "M"))   0.3048)
      ((and (= from-unit "M")   (= to-unit "FT"))  (/ 1.0 0.3048))
      ((and (= from-unit "MM")  (= to-unit "IN"))  (/ 1.0 25.4))
      ((and (= from-unit "IN")  (= to-unit "MM"))  25.4)
      ;; Area
      ((and (= from-unit "SF")  (= to-unit "SY"))  (/ 1.0 9.0))
      ((and (= from-unit "SY")  (= to-unit "SF"))  9.0)
      ((and (= from-unit "SM")  (= to-unit "SY"))  (/ 1.0 0.8361))
      ;; Same unit
      (t 1.0)
    )
  )
  (* value factor)
)

;;;; ----------------------------------------------------------------
;;;; 17. Duplicate-label detection
;;;; ----------------------------------------------------------------

(defun hqto:find-duplicate-labels ( / ss i j ei ej pi pj dups tol)
  "Return a list of pairs (ei ej) of label enames that share the
   same insertion point within tolerance TOL."
  (setq ss   (ssget "X" (list (cons 0 "INSERT")
                              (cons 8 *HQTO-LABEL-LAYER*)
                      )
             )
        tol  1.0
        dups nil
  )
  (if (null ss) (exit))
  (setq i 0)
  (while (< i (1- (sslength ss)))
    (setq ei (ssname ss i)
          pi (cdr (assoc 10 (entget ei)))
          j  (1+ i)
    )
    (while (< j (sslength ss))
      (setq ej (ssname ss j)
            pj (cdr (assoc 10 (entget ej)))
      )
      (if (< (distance pi pj) tol)
        (setq dups (cons (list ei ej) dups))
      )
      (setq j (1+ j))
    )
    (setq i (1+ i))
  )
  dups
)

;;;; ----------------------------------------------------------------
;;;; 18. Selection-set helpers for complex queries
;;;; ----------------------------------------------------------------

(defun hqto:ss-on-layer (layer-name / ss)
  (ssget "X" (list (cons 0 "INSERT") (cons 8 layer-name)))
)

(defun hqto:ss-filter-by-attr (ss tag value / i ename result)
  "Return a new list of enames from SS whose attribute TAG equals VALUE."
  (setq result nil  i 0)
  (while (< i (sslength ss))
    (setq ename (ssname ss i))
    (if (= (strcase (or (hqto:get-attr ename tag) ""))
           (strcase value)
        )
      (setq result (cons ename result))
    )
    (setq i (1+ i))
  )
  result
)

(defun hqto:ss-to-list (ss / i result)
  (setq result nil  i 0)
  (while (< i (sslength ss))
    (setq result (cons (ssname ss i) result))
    (setq i (1+ i))
  )
  result
)

;;;; ----------------------------------------------------------------
;;;; 19. Undo-group wrapper
;;;; ----------------------------------------------------------------

(defun hqto:with-undo-group (func-sym args / res)
  "Execute (apply FUNC-SYM ARGS) inside an UNDO group so the
   entire operation can be reversed with a single U command."
  (command "_.UNDO" "_Begin")
  (setq res (vl-catch-all-apply func-sym args))
  (command "_.UNDO" "_End")
  (if (vl-catch-all-error-p res)
    (progn
      (hqto:log (strcat "Undo-group error: " (vl-catch-all-error-message res)))
      nil
    )
    res
  )
)

;;;; ----------------------------------------------------------------
;;;; 20. Drawing-title-block update helper
;;;; ----------------------------------------------------------------

(defun hqto:update-title-block (dwg-path total-qty unit / ss ename)
  "Write TOTAL-QTY and UNIT into the title-block block (named
   HQTO_TITLEBLOCK) if present in the current drawing."
  (setq ss (ssget "X" '((0 . "INSERT") (2 . "HQTO_TITLEBLOCK"))))
  (if ss
    (progn
      (setq ename (ssname ss 0))
      (hqto:set-attr ename "TOTAL_QTY"  (rtos total-qty 2 4))
      (hqto:set-attr ename "TOTAL_UNIT" unit)
      (hqto:log (strcat "  Title block updated: " (rtos total-qty 2 4) " " unit))
    )
    (hqto:log "  Title block block not found; skipping.")
  )
)

;;;; ----------------------------------------------------------------
;;;; 21. Progressive-save / backup helper
;;;; ----------------------------------------------------------------

(defun hqto:backup-drawing (dwg-path / bak-path)
  "Save a timestamped backup copy of DWG-PATH before processing."
  (setq bak-path
    (strcat (vl-filename-directory dwg-path)
            "\\"
            (vl-filename-base dwg-path)
            "_bak_"
            (rtos (getvar "DATE") 2 0)
            ".dwg"
    )
  )
  (vl-file-copy dwg-path bak-path 1)   ; 1 = overwrite if exists
  (hqto:log (strcat "  Backup saved: " bak-path))
  bak-path
)

;;;; ----------------------------------------------------------------
;;;; 22. Attribute-value validator
;;;; ----------------------------------------------------------------

(defun hqto:validate-label (labelEnt / atts errs cat item unit)
  "Return a list of error strings for LABELENT, or nil if all OK."
  (setq atts (hqto:block-atts labelEnt)
        errs nil
        cat  (cdr (assoc "CATEGORY" atts))
        item (cdr (assoc "ITEM"     atts))
        unit (cdr (assoc "UNIT"     atts))
  )
  (if (or (null cat) (= cat ""))
    (setq errs (cons "Missing CATEGORY" errs))
  )
  (if (or (null item) (= item ""))
    (setq errs (cons "Missing ITEM" errs))
  )
  (if (or (null unit) (= unit ""))
    (setq errs (cons "Missing UNIT" errs))
  )
  errs
)

;;;; ----------------------------------------------------------------
;;;; 23. Batch-validate all labels in current drawing
;;;; ----------------------------------------------------------------

(defun hqto:batch-validate ( / ss i ename errs total-errs)
  (setq ss          (ssget "X" (list (cons 0 "INSERT")
                                     (cons 8 *HQTO-LABEL-LAYER*)
                             )
               )
        total-errs  0
  )
  (if (null ss)
    (progn (hqto:log "Batch-validate: no labels.") (exit))
  )
  (setq i 0)
  (while (< i (sslength ss))
    (setq ename (ssname ss i)
          errs  (hqto:validate-label ename)
    )
    (if errs
      (progn
        (setq total-errs (+ total-errs (length errs)))
        (foreach e errs
          (hqto:log (strcat "  VALIDATE ERR " e
                            " @ " (vl-prin1-to-string ename)
                   )
          )
        )
      )
    )
    (setq i (1+ i))
  )
  total-errs
)

;;;; ----------------------------------------------------------------
;;;; 24. Layer-color synchronisation
;;;; ----------------------------------------------------------------

(defun hqto:sync-layer-colors ()
  "Ensure HQTO layers have the correct standard colors."
  (if (hqto:layer-exists-p *HQTO-LABEL-LAYER*)
    (hqto:make-layer *HQTO-LABEL-LAYER* 3)
  )
  (if (hqto:layer-exists-p *HQTO-RESULT-LAYER*)
    (hqto:make-layer *HQTO-RESULT-LAYER* 1)
  )
)

;;;; ----------------------------------------------------------------
;;;; 25. Freeze / thaw HQTO layers
;;;; ----------------------------------------------------------------

(defun hqto:freeze-hqto-layers ( / cmd-echo)
  (setq cmd-echo (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (command "_.LAYER" "_Freeze" *HQTO-LABEL-LAYER*  ""
                    "_Freeze" *HQTO-RESULT-LAYER* "")
  (setvar "CMDECHO" cmd-echo)
)

(defun hqto:thaw-hqto-layers ( / cmd-echo)
  (setq cmd-echo (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (command "_.LAYER" "_Thaw" *HQTO-LABEL-LAYER*  ""
                    "_Thaw" *HQTO-RESULT-LAYER* "")
  (setvar "CMDECHO" cmd-echo)
)

;;;; ----------------------------------------------------------------
;;;; 26. Export results to Excel via COM (Windows only)
;;;; ----------------------------------------------------------------

(defun hqto:export-to-excel (data-rows out-path / xl wb ws row-idx col-idx row)
  "Write DATA-ROWS (list of lists of strings) to an Excel workbook.
   DATA-ROWS row 0 is treated as the header."
  (setq xl      (vlax-create-object "Excel.Application")
        wb      (vlax-invoke (vlax-get xl 'Workbooks) 'Add)
        ws      (vlax-get (vlax-get wb 'Sheets) 'Item 1)
        row-idx 1
  )
  (vlax-put xl 'Visible 0)
  (foreach row data-rows
    (setq col-idx 1)
    (foreach cell row
      (vlax-put
        (vlax-get
          (vlax-get ws 'Cells)
          'Item row-idx col-idx
        )
        'Value cell
      )
      (setq col-idx (1+ col-idx))
    )
    (setq row-idx (1+ row-idx))
  )
  (vlax-invoke wb 'SaveAs out-path 51)  ; 51 = xlOpenXMLWorkbook
  (vlax-invoke wb 'Close)
  (vlax-invoke xl 'Quit)
  (vlax-release-object ws)
  (vlax-release-object wb)
  (vlax-release-object xl)
  (hqto:log (strcat "Excel export: " out-path))
)

;;;; ----------------------------------------------------------------
;;;; 27. HTML report generator
;;;; ----------------------------------------------------------------

(defun hqto:write-html-report (data-rows out-path / fp row cell first-cell)
  "Write DATA-ROWS as an HTML table to OUT-PATH."
  (setq fp (open out-path "w"))
  (write-line "<!DOCTYPE html><html><head><meta charset='utf-8'>" fp)
  (write-line "<title>Halff QTO Report</title>" fp)
  (write-line "<style>table{border-collapse:collapse}" fp)
  (write-line "td,th{border:1px solid #999;padding:4px 8px}</style>" fp)
  (write-line "</head><body>" fp)
  (write-line (strcat "<h2>Halff QTO Report v" *HQTO-VERSION* "</h2>") fp)
  (write-line "<table>" fp)
  (setq first-cell T)
  (foreach row data-rows
    (write-line (if first-cell "<tr>" "<tr>") fp)
    (setq first-cell nil)
    (foreach cell row
      (write-line (strcat "<td>" cell "</td>") fp)
    )
    (write-line "</tr>" fp)
  )
  (write-line "</table></body></html>" fp)
  (close fp)
  (hqto:log (strcat "HTML report: " out-path))
)

;;;; ----------------------------------------------------------------
;;;; 28. Progress-bar helper (text-window)
;;;; ----------------------------------------------------------------

(defun hqto:progress (current total label / pct bar filled empty-s)
  "Print a simple ASCII progress bar to the text window."
  (setq pct    (if (> total 0) (/ (* current 100.0) total) 0.0)
        filled (fix (/ pct 5))
        empty-s (- 20 filled)
        bar    (strcat "[" (apply 'strcat (mapcar '(lambda (x) "#") (hqto:iota filled)))
                       (apply 'strcat (mapcar '(lambda (x) "-") (hqto:iota empty-s)))
                       "]"
               )
  )
  (princ (strcat "\r" bar " " (rtos pct 2 1) "% " label))
)

(defun hqto:iota (n / result)
  "Return list of N zeros (used as a map target for progress bar)."
  (setq result nil)
  (repeat n (setq result (cons 0 result)))
  result
)

;;;; ----------------------------------------------------------------
;;;; 29. String-padding helpers for aligned text output
;;;; ----------------------------------------------------------------

(defun hqto:pad-right (s width / len)
  "Pad string S with trailing spaces to WIDTH characters."
  (setq len (strlen s))
  (if (>= len width)
    s
    (strcat s (hqto:spaces (- width len)))
  )
)

(defun hqto:pad-left (s width / len)
  "Pad string S with leading spaces to WIDTH characters."
  (setq len (strlen s))
  (if (>= len width)
    s
    (strcat (hqto:spaces (- width len)) s)
  )
)

(defun hqto:spaces (n / result)
  (setq result "")
  (repeat n (setq result (strcat result " ")))
  result
)

;;;; ----------------------------------------------------------------
;;;; 30. Configuration-file reader (INI-style)
;;;; ----------------------------------------------------------------

(defun hqto:read-config (cfg-path / fp line key val config)
  "Parse a simple KEY=VALUE config file.  Returns alist."
  (setq config nil)
  (if (not (findfile cfg-path))
    (progn (hqto:log (strcat "Config not found: " cfg-path)) (exit))
  )
  (setq fp (open cfg-path "r"))
  (while (setq line (read-line fp))
    (setq line (hqto:trim line))
    (if (and (> (strlen line) 0)
             (/= (substr line 1 1) ";")
             (vl-string-search "=" line)
        )
      (progn
        (setq key (hqto:trim (substr line 1 (vl-string-search "=" line)))
              val (hqto:trim (substr line (+ 2 (vl-string-search "=" line))))
        )
        (setq config (cons (cons key val) config))
      )
    )
  )
  (close fp)
  config
)

;;;; ----------------------------------------------------------------
;;;; 31. INI-config writer
;;;; ----------------------------------------------------------------

(defun hqto:write-config (cfg-path config / fp pair)
  "Write ALIST CONFIG to CFG-PATH as KEY=VALUE lines."
  (setq fp (open cfg-path "w"))
  (write-line (strcat "; Halff QTO config  v" *HQTO-VERSION*) fp)
  (write-line (strcat "; Written " (rtos (getvar "DATE") 2 8)) fp)
  (foreach pair config
    (write-line (strcat (car pair) "=" (cdr pair)) fp)
  )
  (close fp)
)

;;;; ----------------------------------------------------------------
;;;; 32. Block-definition census
;;;; ----------------------------------------------------------------

(defun hqto:count-block-defs ( / bt i cnt)
  "Return count of non-anonymous block definitions."
  (setq bt  (vla-get-Blocks (vla-get-ActiveDocument (vlax-get-acad-object)))
        i   0
        cnt 0
  )
  (vlax-for blk bt
    (if (not (= (substr (vlax-get blk 'Name) 1 1) "*"))
      (setq cnt (1+ cnt))
    )
  )
  cnt
)

;;;; ----------------------------------------------------------------
;;;; 33. Attribute-tag renaming utility
;;;; ----------------------------------------------------------------

(defun hqto:rename-attr-tag (blockName oldTag newTag / blkDef attDef atts)
  "Rename attribute definition tag OLDTAG -> NEWTAG in block BLOCKNAME."
  (setq blkDef
    (vla-item
      (vla-get-Blocks (vla-get-ActiveDocument (vlax-get-acad-object)))
      blockName
    )
  )
  (vlax-for obj blkDef
    (if (= (vla-get-ObjectName obj) "AcDbAttributeDefinition")
      (if (= (strcase (vlax-get obj 'TagString)) (strcase oldTag))
        (vlax-put obj 'TagString newTag)
      )
    )
  )
)

;;;; ----------------------------------------------------------------
;;;; 34. Multi-criteria label filter
;;;; ----------------------------------------------------------------

(defun hqto:filter-labels (category item unit / ss result ename atts)
  "Return list of enames matching all non-nil criteria."
  (setq ss     (ssget "X" (list (cons 0 "INSERT")
                                (cons 8 *HQTO-LABEL-LAYER*)
                        )
               )
        result nil
  )
  (if (null ss) (exit))
  (setq i 0)
  (while (< i (sslength ss))
    (setq ename (ssname ss i)
          atts  (hqto:block-atts ename)
    )
    (if (and
          (or (null category)
              (= (strcase (or (cdr (assoc "CATEGORY" atts)) ""))
                 (strcase category)
              )
          )
          (or (null item)
              (= (strcase (or (cdr (assoc "ITEM" atts)) ""))
                 (strcase item)
              )
          )
          (or (null unit)
              (= (strcase (or (cdr (assoc "UNIT" atts)) ""))
                 (strcase unit)
              )
          )
        )
      (setq result (cons ename result))
    )
    (setq i (1+ i))
  )
  result
)

;;;; ----------------------------------------------------------------
;;;; 35. Quantity-rollup by category
;;;; ----------------------------------------------------------------

(defun hqto:rollup-by-category ( / ss i ename atts cat qty rollup entry)
  "Return an alist of (CATEGORY . total-qty) for all labels."
  (setq ss     (ssget "X" (list (cons 0 "INSERT")
                                (cons 8 *HQTO-LABEL-LAYER*)
                        )
               )
        rollup nil
  )
  (if (null ss) (exit))
  (setq i 0)
  (while (< i (sslength ss))
    (setq ename (ssname ss i)
          atts  (hqto:block-atts ename)
          cat   (or (cdr (assoc "CATEGORY" atts)) "UNCATEGORISED")
          qty   (hqto:str->real (or (cdr (assoc "QTY" atts)) "0"))
    )
    (if qty
      (progn
        (setq entry (assoc cat rollup))
        (if entry
          (setcdr entry (+ (cdr entry) qty))
          (setq rollup (cons (cons cat qty) rollup))
        )
      )
    )
    (setq i (1+ i))
  )
  rollup
)

;;;; ----------------------------------------------------------------
;;;; 36. JSON export helper
;;;; ----------------------------------------------------------------

(defun hqto:escape-json-str (s / i c out)
  "Escape special characters for JSON string encoding."
  (setq out ""  i 1)
  (while (<= i (strlen s))
    (setq c (substr s i 1))
    (cond
      ((= c "\\") (setq out (strcat out "\\\\")))
      ((= c "\"") (setq out (strcat out "\\\"")))
      ((= c "\n") (setq out (strcat out "\\n")))
      ((= c "\t") (setq out (strcat out "\\t")))
      (t          (setq out (strcat out c)))
    )
    (setq i (1+ i))
  )
  out
)

(defun hqto:export-json (data-rows out-path / fp row first-row first-cell keys vals)
  "Write DATA-ROWS as a JSON array of objects.  Row 0 = header keys."
  (setq fp        (open out-path "w")
        first-row T
  )
  (write-line "[" fp)
  (setq keys (car data-rows))
  (foreach row (cdr data-rows)
    (if (not first-row) (write-line "," fp))
    (write-line "{" fp)
    (setq first-cell T)
    (mapcar
      '(lambda (k v)
         (if (not first-cell) (write-line "," fp))
         (write-line (strcat "  \"" (hqto:escape-json-str k)
                             "\": \"" (hqto:escape-json-str v) "\""
                    ) fp
         )
         (setq first-cell nil)
       )
      keys row
    )
    (write-line "}" fp)
    (setq first-row nil)
  )
  (write-line "]" fp)
  (close fp)
  (hqto:log (strcat "JSON export: " out-path))
)

;;;; ----------------------------------------------------------------
;;;; 37. Object-snap override helpers
;;;; ----------------------------------------------------------------

(defun hqto:osnap-off ( / )
  (setvar "OSMODE" 0)
)

(defun hqto:osnap-restore (saved-mode)
  (setvar "OSMODE" saved-mode)
)

(defun hqto:with-osnap-off (func-sym args / saved res)
  "Execute FUNC-SYM with ARGS with OSNAP temporarily disabled."
  (setq saved (getvar "OSMODE"))
  (hqto:osnap-off)
  (setq res (vl-catch-all-apply func-sym args))
  (hqto:osnap-restore saved)
  res
)

;;;; ----------------------------------------------------------------
;;;; 38. Drawing-scale reader
;;;; ----------------------------------------------------------------

(defun hqto:get-drawing-scale ( / ann-scale str-scale parts num denom)
  "Return the current annotation scale as a real number (e.g. 0.02083
   for 1:48).  Falls back to 1.0 if unavailable."
  (setq ann-scale
    (vl-catch-all-apply
      '(lambda ()
         (vlax-get
           (vla-get-ActiveDocument (vlax-get-acad-object))
           'AnnotationScaleChanged
         )
       )
    )
  )
  ;; CANNOSCALE sysvar available from AutoCAD 2008+
  (setq str-scale (vl-catch-all-apply 'getvar '("CANNOSCALE")))
  (if (and str-scale (not (vl-catch-all-error-p str-scale))
           (vl-string-search ":" str-scale)
      )
    (progn
      (setq parts  (vl-string-split str-scale ":")
            num    (hqto:str->real (car parts))
            denom  (hqto:str->real (cadr parts))
      )
      (if (and num denom (/= denom 0.0))
        (/ num denom)
        1.0
      )
    )
    1.0
  )
)

;;;; ----------------------------------------------------------------
;;;; 39. Closest-entity finder
;;;; ----------------------------------------------------------------

(defun hqto:nearest-label (pt / ss i ename best-ent best-dist d ctr)
  "Return the label ename whose centre is nearest to PT."
  (setq ss        (ssget "X" (list (cons 0 "INSERT")
                                   (cons 8 *HQTO-LABEL-LAYER*)
                           )
               )
        best-ent  nil
        best-dist 1.0e+30
  )
  (if (null ss) (exit))
  (setq i 0)
  (while (< i (sslength ss))
    (setq ename (ssname ss i)
          ctr   (hqto:ent-center ename)
          d     (distance pt ctr)
    )
    (if (< d best-dist)
      (setq best-dist d  best-ent ename)
    )
    (setq i (1+ i))
  )
  best-ent
)

;;;; ----------------------------------------------------------------
;;;; 40. Attribute quick-print (diagnostic)
;;;; ----------------------------------------------------------------

(defun hqto:print-attrs (labelEnt / atts)
  "Print all attribute tag/value pairs of LABELENT to text window."
  (setq atts (hqto:block-atts labelEnt))
  (princ (strcat "\nAttributes for " (vl-prin1-to-string labelEnt) ":"))
  (foreach pair atts
    (princ (strcat "\n  " (car pair) " = " (cdr pair)))
  )
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 41. Linetype-assignment helper
;;;; ----------------------------------------------------------------

(defun hqto:set-layer-linetype (layer-name lt-name / cmd-echo)
  (setq cmd-echo (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (command "_.LINETYPE" "_Load" lt-name "acad" "")
  (command "_.LAYER" "_LType" lt-name layer-name "")
  (setvar "CMDECHO" cmd-echo)
)

;;;; ----------------------------------------------------------------
;;;; 42. Insertion-point grid helper
;;;; ----------------------------------------------------------------

(defun hqto:grid-points (origin cols rows col-spacing row-spacing / pts r c)
  "Return a flat list of grid insertion points."
  (setq pts nil)
  (setq r 0)
  (while (< r rows)
    (setq c 0)
    (while (< c cols)
      (setq pts
        (cons
          (list (+ (car  origin) (* c col-spacing))
                (+ (cadr origin) (* r row-spacing))
          )
          pts
        )
      )
      (setq c (1+ c))
    )
    (setq r (1+ r))
  )
  (reverse pts)
)

;;;; ----------------------------------------------------------------
;;;; 43. CSV-row builder helper
;;;; ----------------------------------------------------------------

(defun hqto:csv-row (fields / row first-field)
  "Join FIELDS list into a CSV row, quoting fields containing commas."
  (setq row         ""
        first-field T
  )
  (foreach f fields
    (if (not first-field) (setq row (strcat row ",")))
    (if (vl-string-search "," f)
      (setq row (strcat row "\"" f "\""))
      (setq row (strcat row f))
    )
    (setq first-field nil)
  )
  row
)

;;;; ----------------------------------------------------------------
;;;; 44. Interactive region-draw helper
;;;; ----------------------------------------------------------------

(defun hqto:user-draw-region ( / pt1 pt2 pts)
  "Prompt user to pick two corners of a rectangular region.
   Returns list of 4 2-D points, or nil if cancelled."
  (setq pt1 (getpoint "\nPick first corner of region: "))
  (if (null pt1) nil
    (progn
      (setq pt2 (getcorner pt1 "\nPick opposite corner: "))
      (if (null pt2) nil
        (list
          (list (car pt1) (cadr pt1))
          (list (car pt2) (cadr pt1))
          (list (car pt2) (cadr pt2))
          (list (car pt1) (cadr pt2))
        )
      )
    )
  )
)

;;;; ----------------------------------------------------------------
;;;; 45. Annotation-scale-aware text height
;;;; ----------------------------------------------------------------

(defun hqto:scaled-text-height (base-height)
  "Return BASE-HEIGHT adjusted for the current annotation scale."
  (/ base-height (hqto:get-drawing-scale))
)

;;;; ----------------------------------------------------------------
;;;; 46. System-variable save/restore stack
;;;; ----------------------------------------------------------------

(setq *HQTO-SYSVAR-STACK* nil)

(defun hqto:push-sysvar (var-name)
  (setq *HQTO-SYSVAR-STACK*
    (cons (cons var-name (getvar var-name))
          *HQTO-SYSVAR-STACK*
    )
  )
)

(defun hqto:pop-sysvar ( / top)
  (if *HQTO-SYSVAR-STACK*
    (progn
      (setq top (car *HQTO-SYSVAR-STACK*))
      (setvar (car top) (cdr top))
      (setq *HQTO-SYSVAR-STACK* (cdr *HQTO-SYSVAR-STACK*))
    )
  )
)

(defun hqto:pop-all-sysvars ()
  (while *HQTO-SYSVAR-STACK*
    (hqto:pop-sysvar)
  )
)

;;;; ----------------------------------------------------------------
;;;; 47. Reactor – auto-update label on attribute change
;;;; ----------------------------------------------------------------

(defun hqto:install-change-reactor ( / r)
  "Install an object reactor so that editing a label attribute
   triggers an automatic quantity re-compute."
  (setq r
    (vlr-object-reactor
      nil  ; no trigger object list at install time
      "HQTOAttrReactor"
      '((:vlr-modified . hqto:reactor-modified-cb))
    )
  )
  r
)

(defun hqto:reactor-modified-cb (notifier-obj reactor-obj param-list)
  "Reactor callback: re-compute qty when a label attribute is modified."
  (vl-catch-all-apply
    '(lambda ()
       (setq ename (vlax-vla-object->ename notifier-obj))
       (if (and ename
                (= (cdr (assoc 8 (entget ename))) *HQTO-LABEL-LAYER*)
           )
         (progn
           (setq qty (hqto:compute-qty ename))
           (hqto:set-attr ename "QTY" (rtos qty 2 4))
         )
       )
     )
  )
)

;;;; ----------------------------------------------------------------
;;;; 48. Named UCS helper
;;;; ----------------------------------------------------------------

(defun hqto:save-ucs (ucs-name / cmd-echo)
  (setq cmd-echo (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (command "_.UCS" "_Save" ucs-name)
  (setvar "CMDECHO" cmd-echo)
)

(defun hqto:restore-ucs (ucs-name / cmd-echo)
  (setq cmd-echo (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (command "_.UCS" "_Restore" ucs-name)
  (setvar "CMDECHO" cmd-echo)
)

;;;; ----------------------------------------------------------------
;;;; 49. Clipboard export (Windows)
;;;; ----------------------------------------------------------------

(defun hqto:copy-to-clipboard (text / wsh)
  "Place TEXT on the Windows clipboard using WScript.Shell."
  (setq wsh (vlax-create-object "WScript.Shell"))
  (vl-catch-all-apply
    '(lambda ()
       (vlax-invoke wsh 'Run
         (strcat "cmd /c echo " text " | clip")
         0 T
       )
     )
  )
  (vlax-release-object wsh)
)

;;;; ----------------------------------------------------------------
;;;; 50. End-of-file / load confirmation
;;;; ----------------------------------------------------------------

(princ
  (strcat
    "\nHalff QTO Labels v" *HQTO-VERSION* " loaded."
    "\n  Commands: QLABEL  QRUN  QREPORT  QCHECK"
    "\n           QFAILSEARCH"
    "\n  TEXT QTY extraction supported (QTY_TYPE=TEXT attribute)."
  )
)
(princ "\n")

;; ---- extended help / command list (sections 51-100 follow) --------

;;;; ----------------------------------------------------------------
;;;; 51. QHELP command
;;;; ----------------------------------------------------------------

(defun C:QHELP ( / )
  (princ "\n========================================")
  (princ (strcat "\n  Halff QTO Labels v" *HQTO-VERSION*))
  (princ "\n========================================")
  (princ "\n")
  (princ "\n  QLABEL       - Place a quantity label")
  (princ "\n  QRUN         - Batch-process drawings")
  (princ "\n  QREPORT      - Write CSV summary")
  (princ "\n  QCHECK       - Audit labels in drawing")
  (princ "\n  QFAILSEARCH  - Step through failures")
  (princ "\n  QHELP        - Show this help")
  (princ "\n  QSTATS       - Show session statistics")
  (princ "\n  QCLEAR       - Clear session data")
  (princ "\n  QCONFIG      - Edit configuration")
  (princ "\n  QEXCEL       - Export results to Excel")
  (princ "\n  QJSON        - Export results to JSON")
  (princ "\n  QHTML        - Export results to HTML")
  (princ "\n  QZOOM        - Zoom to label")
  (princ "\n  QDUPLICATE   - Find duplicate labels")
  (princ "\n  QROLLUP      - Show category rollup")
  (princ "\n  QVALIDATE    - Validate all labels")
  (princ "\n  QREACTOR     - Install change reactor")
  (princ "\n")
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 52. QSTATS command
;;;; ----------------------------------------------------------------

(defun C:QSTATS ( / ss cnt-label cnt-fail)
  (setq ss        (ssget "X" (list (cons 0 "INSERT")
                                   (cons 8 *HQTO-LABEL-LAYER*)
                           )
               )
        cnt-label (if ss (sslength ss) 0)
        cnt-fail  (length *HQTO-FAIL-LIST*)
  )
  (princ (strcat "\n===== HQTO Session Statistics ====="))
  (princ (strcat "\n  Labels in current drawing : " (itoa cnt-label)))
  (princ (strcat "\n  Failures this session     : " (itoa cnt-fail)))
  (princ (strcat "\n  DWGs processed            : "
                 (itoa (length *HQTO-DWG-LIST*))
         )
  )
  (princ (strcat "\n  Log file                  : " *HQTO-LOG-FILE*))
  (princ "\n")
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 53. QCLEAR command
;;;; ----------------------------------------------------------------

(defun C:QCLEAR ( / )
  (setq *HQTO-FAIL-LIST* nil
        *HQTO-FAIL-IDX*  0
        *HQTO-DWG-LIST*  nil
  )
  (princ "\nSession data cleared.")
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 54. QCONFIG command
;;;; ----------------------------------------------------------------

(defun C:QCONFIG ( / cfg-path config new-val)
  (setq cfg-path
    (getfiled "Select config file" ""
              "ini" 2)
  )
  (if (null cfg-path)
    (progn (princ "\nQCONFIG cancelled.") (princ) (exit))
  )
  (setq config (hqto:read-config cfg-path))
  (princ (strcat "\nLoaded " (itoa (length config)) " settings from " cfg-path))
  (foreach pair config
    (princ (strcat "\n  " (car pair) " = " (cdr pair)))
  )
  ;; Allow override of LABEL_LAYER
  (setq new-val
    (getstring T (strcat "\nNew LABEL_LAYER [" *HQTO-LABEL-LAYER* "]: "))
  )
  (if (and new-val (/= new-val ""))
    (setq *HQTO-LABEL-LAYER* new-val)
  )
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 55. QEXCEL command
;;;; ----------------------------------------------------------------

(defun C:QEXCEL ( / out-path data-rows ss i ename atts row)
  (setq out-path
    (getfiled "Save Excel report as" ""
              "xlsx" 1)
  )
  (if (null out-path)
    (progn (princ "\nQEXCEL cancelled.") (princ) (exit))
  )
  (setq data-rows
    (list '("Drawing" "Category" "Item" "QTY" "Unit" "Note"))
  )
  (setq ss (ssget "X" (list (cons 0 "INSERT")
                            (cons 8 *HQTO-LABEL-LAYER*)
                    )
           )
  )
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ename (ssname ss i)
              atts  (hqto:block-atts ename)
        )
        (setq row
          (list
            (getvar "DWGNAME")
            (or (cdr (assoc "CATEGORY" atts)) "")
            (or (cdr (assoc "ITEM"     atts)) "")
            (or (cdr (assoc "QTY"      atts)) "0")
            (or (cdr (assoc "UNIT"     atts)) "")
            (or (cdr (assoc "NOTE"     atts)) "")
          )
        )
        (setq data-rows (append data-rows (list row)))
        (setq i (1+ i))
      )
    )
  )
  (hqto:export-to-excel data-rows out-path)
  (princ (strcat "\nExcel report written: " out-path))
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 56. QJSON command
;;;; ----------------------------------------------------------------

(defun C:QJSON ( / out-path data-rows ss i ename atts row)
  (setq out-path
    (getfiled "Save JSON report as" ""
              "json" 1)
  )
  (if (null out-path)
    (progn (princ "\nQJSON cancelled.") (princ) (exit))
  )
  (setq data-rows
    (list '("Drawing" "Category" "Item" "QTY" "Unit" "Note"))
  )
  (setq ss (ssget "X" (list (cons 0 "INSERT")
                            (cons 8 *HQTO-LABEL-LAYER*)
                    )
           )
  )
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ename (ssname ss i)
              atts  (hqto:block-atts ename)
        )
        (setq row
          (list
            (getvar "DWGNAME")
            (or (cdr (assoc "CATEGORY" atts)) "")
            (or (cdr (assoc "ITEM"     atts)) "")
            (or (cdr (assoc "QTY"      atts)) "0")
            (or (cdr (assoc "UNIT"     atts)) "")
            (or (cdr (assoc "NOTE"     atts)) "")
          )
        )
        (setq data-rows (append data-rows (list row)))
        (setq i (1+ i))
      )
    )
  )
  (hqto:export-json data-rows out-path)
  (princ (strcat "\nJSON report written: " out-path))
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 57. QHTML command
;;;; ----------------------------------------------------------------

(defun C:QHTML ( / out-path data-rows ss i ename atts row)
  (setq out-path
    (getfiled "Save HTML report as" ""
              "html" 1)
  )
  (if (null out-path)
    (progn (princ "\nQHTML cancelled.") (princ) (exit))
  )
  (setq data-rows
    (list '("Drawing" "Category" "Item" "QTY" "Unit" "Note"))
  )
  (setq ss (ssget "X" (list (cons 0 "INSERT")
                            (cons 8 *HQTO-LABEL-LAYER*)
                    )
           )
  )
  (if ss
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ename (ssname ss i)
              atts  (hqto:block-atts ename)
        )
        (setq row
          (list
            (getvar "DWGNAME")
            (or (cdr (assoc "CATEGORY" atts)) "")
            (or (cdr (assoc "ITEM"     atts)) "")
            (or (cdr (assoc "QTY"      atts)) "0")
            (or (cdr (assoc "UNIT"     atts)) "")
            (or (cdr (assoc "NOTE"     atts)) "")
          )
        )
        (setq data-rows (append data-rows (list row)))
        (setq i (1+ i))
      )
    )
  )
  (hqto:write-html-report data-rows out-path)
  (princ (strcat "\nHTML report written: " out-path))
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 58. QZOOM command  -- zoom to selected label
;;;; ----------------------------------------------------------------

(defun C:QZOOM ( / pt ename)
  (setq pt (getpoint "\nPick near a label to zoom to: "))
  (if pt
    (progn
      (setq ename (hqto:nearest-label pt))
      (if ename
        (command "_.ZOOM" "_Object" ename "")
        (princ "\nNo label found.")
      )
    )
  )
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 59. QDUPLICATE command
;;;; ----------------------------------------------------------------

(defun C:QDUPLICATE ( / dups)
  (setq dups (hqto:find-duplicate-labels))
  (if (null dups)
    (princ "\nNo duplicate labels found.")
    (progn
      (princ (strcat "\n" (itoa (length dups)) " duplicate pair(s):"))
      (foreach pair dups
        (princ (strcat "\n  " (vl-prin1-to-string (car pair))
                       " <-> " (vl-prin1-to-string (cadr pair))
               )
        )
      )
    )
  )
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 60. QROLLUP command
;;;; ----------------------------------------------------------------

(defun C:QROLLUP ( / rollup)
  (setq rollup (hqto:rollup-by-category))
  (princ "\n===== QTO Rollup by Category =====")
  (foreach pair rollup
    (princ (strcat "\n  "
                   (hqto:pad-right (car pair) 24)
                   (hqto:pad-left  (rtos (cdr pair) 2 4) 12)
           )
    )
  )
  (princ "\n")
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 61. QVALIDATE command
;;;; ----------------------------------------------------------------

(defun C:QVALIDATE ( / n)
  (setq n (hqto:batch-validate))
  (princ (strcat "\nQVALIDATE: " (itoa n) " error(s) found."))
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 62. QREACTOR command
;;;; ----------------------------------------------------------------

(defun C:QREACTOR ( / r)
  (setq r (hqto:install-change-reactor))
  (if r
    (princ "\nChange reactor installed.")
    (princ "\nFailed to install reactor.")
  )
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 63-80. Placeholder sections for future extension
;;;;         (reserved for v2.1 features)
;;;; ----------------------------------------------------------------

;; Section 63 - reserved
;; Section 64 - reserved
;; Section 65 - reserved
;; Section 66 - reserved
;; Section 67 - reserved
;; Section 68 - reserved
;; Section 69 - reserved
;; Section 70 - reserved
;; Section 71 - reserved
;; Section 72 - reserved
;; Section 73 - reserved
;; Section 74 - reserved
;; Section 75 - reserved
;; Section 76 - reserved
;; Section 77 - reserved
;; Section 78 - reserved
;; Section 79 - reserved
;; Section 80 - reserved

;;;; ----------------------------------------------------------------
;;;; 81. String-split utility (for config reader)
;;;; ----------------------------------------------------------------

(defun vl-string-split (str delim / result start i c sub)
  "Split STR by single-character string DELIM.  Returns list of parts."
  (setq result nil
        start  1
        i      1
  )
  (while (<= i (strlen str))
    (setq c (substr str i 1))
    (if (= c delim)
      (progn
        (setq sub (substr str start (- i start)))
        (setq result (append result (list sub)))
        (setq start (1+ i))
      )
    )
    (setq i (1+ i))
  )
  ;; last segment
  (setq result (append result (list (substr str start))))
  result
)

;;;; ----------------------------------------------------------------
;;;; 82. List-dedup utility
;;;; ----------------------------------------------------------------

(defun hqto:list-dedup (lst / seen result)
  "Return LST with duplicate elements removed (preserves order)."
  (setq seen nil  result nil)
  (foreach x lst
    (if (not (member x seen))
      (progn
        (setq seen   (cons x seen)
              result (cons x result)
        )
      )
    )
  )
  (reverse result)
)

;;;; ----------------------------------------------------------------
;;;; 83. List-sort wrapper (insertion sort for short lists)
;;;; ----------------------------------------------------------------

(defun hqto:sort-strings (lst / sorted insert-sorted)
  "Sort list of strings alphabetically (ascending)."
  (defun insert-sorted (s sorted)
    (cond
      ((null sorted) (list s))
      ((<= (ascii s) (ascii (car sorted)))
       (cons s sorted)
      )
      (t (cons (car sorted) (insert-sorted s (cdr sorted))))
    )
  )
  (setq sorted nil)
  (foreach s lst
    (setq sorted (insert-sorted s sorted))
  )
  sorted
)

;;;; ----------------------------------------------------------------
;;;; 84. Numeric-string right-align for tables
;;;; ----------------------------------------------------------------

(defun hqto:fmt-qty (qty-real decimals width)
  "Format QTY-REAL to DECIMALS decimal places, right-padded to WIDTH."
  (hqto:pad-left (rtos qty-real 2 decimals) width)
)

;;;; ----------------------------------------------------------------
;;;; 85. Attribute-definition existence check
;;;; ----------------------------------------------------------------

(defun hqto:block-has-attr-p (blk-name tag / blkDef found)
  "Return T if block definition BLK-NAME contains attribute tag TAG."
  (setq blkDef
    (vl-catch-all-apply
      '(lambda ()
         (vla-item
           (vla-get-Blocks (vla-get-ActiveDocument (vlax-get-acad-object)))
           blk-name
         )
       )
    )
  )
  (if (vl-catch-all-error-p blkDef)
    nil
    (progn
      (setq found nil)
      (vlax-for obj blkDef
        (if (and (= (vla-get-ObjectName obj) "AcDbAttributeDefinition")
                 (= (strcase (vlax-get obj 'TagString)) (strcase tag))
            )
          (setq found T)
        )
      )
      found
    )
  )
)

;;;; ----------------------------------------------------------------
;;;; 86. Drawing-list persistence (save/load between sessions)
;;;; ----------------------------------------------------------------

(defun hqto:save-dwg-list (path / fp)
  (setq fp (open path "w"))
  (foreach dwg *HQTO-DWG-LIST*
    (write-line dwg fp)
  )
  (close fp)
)

(defun hqto:load-dwg-list (path / fp line lst)
  (setq fp (open path "r")
        lst nil
  )
  (while (setq line (read-line fp))
    (setq lst (cons (hqto:trim line) lst))
  )
  (close fp)
  (setq *HQTO-DWG-LIST* (reverse lst))
)

;;;; ----------------------------------------------------------------
;;;; 87. Failure-list export
;;;; ----------------------------------------------------------------

(defun hqto:export-failures (out-path / fp entry)
  "Write *HQTO-FAIL-LIST* to OUT-PATH as a CSV."
  (setq fp (open out-path "w"))
  (write-line "Drawing,Entity,Message" fp)
  (foreach entry *HQTO-FAIL-LIST*
    (write-line
      (hqto:csv-row
        (list
          (or (car   entry) "")
          (vl-prin1-to-string (or (cadr  entry) ""))
          (or (caddr entry) "")
        )
      )
      fp
    )
  )
  (close fp)
)

;;;; ----------------------------------------------------------------
;;;; 88. Batch-backup before QRUN
;;;; ----------------------------------------------------------------

(defun hqto:batch-backup (dwg-dir / file-list)
  "Create timestamped backups of all DWGs in DWG-DIR."
  (setq file-list (vl-directory-files dwg-dir "*.dwg" 1))
  (foreach dwg file-list
    (hqto:backup-drawing (strcat dwg-dir "\\" dwg))
  )
)

;;;; ----------------------------------------------------------------
;;;; 89. QBACKUP command
;;;; ----------------------------------------------------------------

(defun C:QBACKUP ( / dwg-dir)
  (setq dwg-dir
    (vl-filename-directory
      (getfiled "Select any DWG in folder to backup" "" "dwg" 2)
    )
  )
  (if (null dwg-dir)
    (progn (princ "\nQBACKUP cancelled.") (princ) (exit))
  )
  (hqto:batch-backup dwg-dir)
  (princ "\nBackup complete.")
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 90. QFAILEXPORT command
;;;; ----------------------------------------------------------------

(defun C:QFAILEXPORT ( / out-path)
  (if (null *HQTO-FAIL-LIST*)
    (progn (princ "\nNo failures to export.") (princ) (exit))
  )
  (setq out-path
    (getfiled "Save failure report as" ""
              "csv" 1)
  )
  (if (null out-path)
    (progn (princ "\nCancelled.") (princ) (exit))
  )
  (hqto:export-failures out-path)
  (princ (strcat "\nFailures exported to " out-path))
  (princ)
)

;;;; ----------------------------------------------------------------
;;;; 91. Object-count helper
;;;; ----------------------------------------------------------------

(defun hqto:count-objects-in-region (region-pts ent-type / ss)
  "Return count of ENT-TYPE entities inside REGION-PTS polygon."
  (setq ss (ssget "CP" region-pts (list (cons 0 ent-type))))
  (if ss (sslength ss) 0)
)

;;;; ----------------------------------------------------------------
;;;; 92. Layer-listing helper
;;;; ----------------------------------------------------------------

(defun hqto:list-layers ( / result)
  "Return a list of all layer names in the current drawing."
  (setq result nil)
  (vlax-for lyr
    (vla-get-Layers (vla-get-ActiveDocument (vlax-get-acad-object)))
    (setq result (cons (vlax-get lyr 'Name) result))
  )
  (hqto:sort-strings result)
)

;;;; ----------------------------------------------------------------
;;;; 93. Block-name listing helper
;;;; ----------------------------------------------------------------

(defun hqto:list-block-names ( / result)
  "Return sorted list of non-anonymous block definition names."
  (setq result nil)
  (vlax-for blk
    (vla-get-Blocks (vla-get-ActiveDocument (vlax-get-acad-object)))
    (if (not (= (substr (vlax-get blk 'Name) 1 1) "*"))
      (setq result (cons (vlax-get blk 'Name) result))
    )
  )
  (hqto:sort-strings result)
)

;;;; ----------------------------------------------------------------
;;;; 94. Zoom-to-extents helper
;;;; ----------------------------------------------------------------

(defun hqto:zoom-extents ( / cmd-echo)
  (setq cmd-echo (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (command "_.ZOOM" "_Extents")
  (setvar "CMDECHO" cmd-echo)
)

;;;; ----------------------------------------------------------------
;;;; 95. Purge-unused HQTO layers
;;;; ----------------------------------------------------------------

(defun hqto:purge-hqto-layers ( / cmd-echo)
  (setq cmd-echo (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (command "_.PURGE" "_LAyers" *HQTO-LABEL-LAYER*  "N"
           "_.PURGE" "_LAyers" *HQTO-RESULT-LAYER* "N"
  )
  (setvar "CMDECHO" cmd-echo)
)

;;;; ----------------------------------------------------------------
;;;; 96. REGEN wrapper
;;;; ----------------------------------------------------------------

(defun hqto:regen ( / cmd-echo)
  (setq cmd-echo (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (command "_.REGEN")
  (setvar "CMDECHO" cmd-echo)
)

;;;; ----------------------------------------------------------------
;;;; 97. Attribute-value copy between entities
;;;; ----------------------------------------------------------------

(defun hqto:copy-attrs (src-ent dst-ent / src-atts)
  "Copy all attribute values from SRC-ENT to DST-ENT (matching tags)."
  (setq src-atts (hqto:block-atts src-ent))
  (foreach pair src-atts
    (hqto:set-attr dst-ent (car pair) (cdr pair))
  )
)

;;;; ----------------------------------------------------------------
;;;; 98. Text-height normaliser
;;;; ----------------------------------------------------------------

(defun hqto:normalize-text-height (ss target-height / i ename ed)
  "Set the text height of all TEXT entities in SS to TARGET-HEIGHT."
  (setq i 0)
  (while (< i (sslength ss))
    (setq ename (ssname ss i)
          ed    (entget ename)
    )
    (if (= (cdr (assoc 0 ed)) "TEXT")
      (entmod (subst (cons 40 target-height)
                     (assoc 40 ed)
                     ed
              )
      )
    )
    (setq i (1+ i))
  )
)

;;;; ----------------------------------------------------------------
;;;; 99. Entity-type census
;;;; ----------------------------------------------------------------

(defun hqto:entity-type-census ( / ss i ename etype census entry)
  "Return alist of (ENTITY-TYPE . count) for all entities in drawing."
  (setq ss     (ssget "X")
        census nil
  )
  (if (null ss) (exit))
  (setq i 0)
  (while (< i (sslength ss))
    (setq ename (ssname ss i)
          etype (cdr (assoc 0 (entget ename)))
          entry (assoc etype census)
    )
    (if entry
      (setcdr entry (1+ (cdr entry)))
      (setq census (cons (cons etype 1) census))
    )
    (setq i (1+ i))
  )
  census
)

;;;; ----------------------------------------------------------------
;;;; 100. Final load message & command summary
;;;; ----------------------------------------------------------------

(princ "\n--- Halff QTO Labels v2.0 fully loaded ---")
(princ "\n  Full command list: QHELP")
(princ (strcat "\n  QLABEL  QRUN  QREPORT  QCHECK"))
(princ (strcat "\n  QSTATS  QCLEAR  QCONFIG  QBACKUP  QFAILEXPORT"))
(princ (strcat "\n  QEXCEL  QJSON  QHTML  QZOOM  QDUPLICATE"))
(princ (strcat "\n  QROLLUP  QVALIDATE  QREACTOR  QHELP"))
(princ (strcat "\n  QFAILSEARCH     - Step through failure entities (QRUN)"))
(princ "\n")
(princ)