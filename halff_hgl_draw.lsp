;; =====================================================================
;; HGL Profile Polyline - halff_hgl_draw.lsp
;;
;; Excel columns (row 1 = header, data starts row 2):
;;   A: Pipe
;;   B: Alignment
;;   C: Profile View Name
;;   D: DS Station
;;   E: US Station
;;   F: DS Design HGL  (optional; drawn as a 3-ft horizontal stub)
;;   G: DS HGL
;;   H: US HGL
;;   I: US Design HGL  (optional; drawn as a 3-ft horizontal stub)
;;
;; HGLDRAW processes ALL profile views in the Excel in one run.
;; Rows are grouped by Profile View Name (col C).  For each group the
;; routine tries to locate the matching AECC_PROFILE_VIEW entity in the
;; drawing by name and auto-read its parameters; falling back to manual
;; entry if the entity is not found or COM read fails.  Parameters
;; (scales, datums) carry over as defaults between views so you only
;; need to re-type values that differ.
;;
;; One HGL polyline is drawn per profile view on the CURRENT active layer.
;; No new layers are created.
;;
;; Design HGL stubs: for every non-empty DS/US Design HGL cell a
;; separate 2-vertex polyline is drawn, centered on the pipe's DS or US
;; station at the design elevation, extending 1.5 station-feet each
;; side (3 ft total).  These are left for manual adjustment.
;;
;; Commands:
;;   HGLSET  - browse to and save the Excel file path
;;   HGLDRAW - read Excel, draw all HGL polylines and Design HGL stubs
;; =====================================================================

(vl-load-com)

;; -------------------------------------------------------------------
;; UTILITY
;; -------------------------------------------------------------------

(defun hgl:variant->val (v)
  (cond
    ((= (type v) 'VARIANT)   (hgl:variant->val (vlax-variant-value v)))
    ((= (type v) 'SAFEARRAY) (vlax-safearray->list v))
    (T v)))

(defun hgl:trim (s)
  (if (and s (= (type s) 'STR))
    (vl-string-trim " \t\r\n" s)
    (if s (vl-princ-to-string s) "")))

(defun hgl:num-p (v)
  (or (= (type v) 'REAL) (= (type v) 'INT)))

;; Try each property name in order; return first non-error value or nil.
(defun hgl:try-prop (obj names / res)
  (setq res nil)
  (while (and names (not res))
    (setq res (vl-catch-all-apply 'vlax-get-property (list obj (car names))))
    (if (vl-catch-all-error-p res) (setq res nil))
    (setq names (cdr names)))
  res)

;; -------------------------------------------------------------------
;; SIDECAR - save/load Excel path alongside the DWG
;; -------------------------------------------------------------------

(defun hgl:sidecar-path ()
  (strcat (getvar "DWGPREFIX") (getvar "DWGNAME") ".hglpath"))

(defun hgl:get-excel-path (/ f ln)
  (if (findfile (hgl:sidecar-path))
    (progn
      (setq f (open (hgl:sidecar-path) "r"))
      (setq ln (if f (read-line f) nil))
      (if f (close f))
      ln)
    nil))

(defun hgl:set-excel-path (p / f)
  (setq f (open (hgl:sidecar-path) "w"))
  (if f (progn (write-line p f) (close f) T) nil))

;; -------------------------------------------------------------------
;; EXCEL OPEN / CLOSE
;; -------------------------------------------------------------------

(defun hgl:xl-open (xlsx / xl wbs wb res)
  (setq res (vl-catch-all-apply
    '(lambda ()
       (setq xl (vlax-get-or-create-object "Excel.Application"))
       (vlax-put-property xl 'Visible :vlax-false)
       (vlax-put-property xl 'DisplayAlerts :vlax-false)
       (vlax-put-property xl 'ScreenUpdating :vlax-false)
       (setq wbs (vlax-get-property xl 'Workbooks))
       (setq wb  (vlax-invoke-method wbs 'Open xlsx))
       (list xl wb))))
  (if (vl-catch-all-error-p res)
    (progn (princ "\nERROR: Could not open Excel file.") nil)
    res))

(defun hgl:xl-close (xl wb)
  (if wb (vl-catch-all-apply 'vlax-invoke-method (list wb 'Close 0)))
  (if xl (progn
           (vl-catch-all-apply 'vlax-invoke-method (list xl 'Quit))
           (vl-catch-all-apply 'vlax-release-object (list xl)))))

;; -------------------------------------------------------------------
;; CELL READING
;; -------------------------------------------------------------------

(defun hgl:get-cell (ws row col / cells rng val)
  (setq cells (vlax-get-property ws 'Cells))
  (setq rng   (vlax-get-property cells 'Item row col))
  (if (= (type rng) 'VARIANT) (setq rng (vlax-variant-value rng)))
  (setq val (vlax-get-property rng 'Value))
  (if (= (type val) 'VARIANT) (vlax-variant-value val) val))

(defun hgl:cell-str (ws row col / res)
  (setq res (vl-catch-all-apply 'hgl:get-cell (list ws row col)))
  (if (vl-catch-all-error-p res) ""
    (hgl:trim (vl-princ-to-string (hgl:variant->val res)))))

(defun hgl:cell-num (ws row col / res v)
  (setq res (vl-catch-all-apply 'hgl:get-cell (list ws row col)))
  (if (vl-catch-all-error-p res) nil
    (progn
      (setq v (hgl:variant->val res))
      (cond
        ((= (type v) 'REAL) v)
        ((= (type v) 'INT)  (float v))
        ((and (= (type v) 'STR) (/= v "")) (atof v))
        (T nil)))))

(defun hgl:used-rows (ws / ur)
  (setq ur (vlax-get-property ws 'UsedRange))
  (vlax-get-property (vlax-get-property ur 'Rows) 'Count))

;; -------------------------------------------------------------------
;; READ HGL DATA FROM EXCEL
;;
;; Returns a list of rows:
;;   (pipe align pv-name ds-sta us-sta ds-design ds-hgl us-hgl us-design)
;;   Index:  0     1      2       3       4         5        6      7      8
;;
;; ds-design (col F) and us-design (col I) may be nil.
;; Stops on the first row with an empty DS Station cell (col D).
;; -------------------------------------------------------------------

(defun hgl:read-excel (xlsx / pair xl wb ws nrows row data
                             pipe align pv ds-sta us-sta ds-design ds us us-design)
  (setq data nil)
  (setq pair (hgl:xl-open xlsx))
  (if (not pair)
    nil
    (progn
      (setq xl (car pair) wb (cadr pair))
      (setq ws (vlax-get-property wb 'ActiveSheet))
      (setq nrows (hgl:used-rows ws))
      (setq row 2)
      (while (<= row nrows)
        (setq ds-sta (hgl:cell-num ws row 4))  ;; col D: DS Station
        (if (not ds-sta)
          (setq row (1+ nrows))
          (progn
            (setq pipe      (hgl:cell-str ws row 1)   ;; A: Pipe
                  align     (hgl:cell-str ws row 2)   ;; B: Alignment
                  pv        (hgl:cell-str ws row 3)   ;; C: Profile View Name
                  us-sta    (hgl:cell-num ws row 5)   ;; E: US Station
                  ds-design (hgl:cell-num ws row 6)   ;; F: DS Design HGL
                  ds        (hgl:cell-num ws row 7)   ;; G: DS HGL
                  us        (hgl:cell-num ws row 8)   ;; H: US HGL
                  us-design (hgl:cell-num ws row 9))  ;; I: US Design HGL
            (if (and us-sta ds us)
              (setq data (append data
                (list (list pipe align pv ds-sta us-sta ds-design ds us us-design)))))
            (setq row (1+ row)))))
      (hgl:xl-close xl wb)
      data)))

;; -------------------------------------------------------------------
;; GET UNIQUE PROFILE VIEW NAMES in order of first appearance
;; -------------------------------------------------------------------

(defun hgl:unique-pv-names (data / names drow pv)
  (setq names '())
  (foreach drow data
    (setq pv (hgl:trim (nth 2 drow)))
    (if (and (/= pv "") (not (member pv names)))
      (setq names (append names (list pv)))))
  names)

;; -------------------------------------------------------------------
;; FIND AECC_PROFILE_VIEW ENTITY BY NAME (case-insensitive)
;; Returns the ename or nil.
;; -------------------------------------------------------------------

(defun hgl:find-pv-by-name (pv-name / ss idx ent vla name-val found)
  (setq ss (ssget "X" '((0 . "AECC_PROFILE_VIEW"))))
  (setq found nil)
  (if ss
    (progn
      (setq idx 0)
      (while (and (< idx (sslength ss)) (not found))
        (setq ent (ssname ss idx))
        (setq vla (vl-catch-all-apply 'vlax-ename->vla-object (list ent)))
        (if (not (vl-catch-all-error-p vla))
          (progn
            (setq name-val (vl-catch-all-apply 'vlax-get-property (list vla 'Name)))
            (if (and (not (vl-catch-all-error-p name-val))
                     (= (strcase (hgl:trim (vl-princ-to-string name-val)))
                        (strcase pv-name)))
              (setq found ent))))
        (setq idx (1+ idx)))))
  found)

;; -------------------------------------------------------------------
;; CIVIL 3D PROFILE VIEW AUTO-READ
;;
;; Reads origin, datum station, datum elevation, scales, and direction
;; from an AECC_PROFILE_VIEW entity via COM.
;;
;; SwapedViewDirection (Civil 3D's own typo) = T means right-to-left.
;; h-scale is returned negative for R-L profiles.
;;
;; Returns (ox oy sta-datum elev-datum h-scale v-scale is-rl) or nil.
;; -------------------------------------------------------------------

(defun hgl:pv-read (ent / vla loc ox oy sta-start elev-min h-raw v-raw rl-raw is-rl res)
  (if (/= (cdr (assoc 0 (entget ent))) "AECC_PROFILE_VIEW")
    (progn (princ "\n  Not an AECC_PROFILE_VIEW.") nil)
    (progn
      (setq res (vl-catch-all-apply
        '(lambda ()
           (setq vla (vlax-ename->vla-object ent))

           (setq loc (hgl:try-prop vla '(Location InsertionPoint Origin)))
           (if (= (type loc) 'VARIANT) (setq loc (vlax-variant-value loc)))
           (cond
             ((= (type loc) 'SAFEARRAY)
              (setq ox (vlax-safearray-get-element loc 0)
                    oy (vlax-safearray-get-element loc 1)))
             ((listp loc) (setq ox (car loc) oy (cadr loc)))
             (T (setq ox nil oy nil)))

           (setq sta-start (hgl:try-prop vla '(StationStart StartStation)))
           (if (and sta-start (= (type sta-start) 'VARIANT))
             (setq sta-start (vlax-variant-value sta-start)))

           (setq elev-min (hgl:try-prop vla '(ElevationMin MinimumElevation DatumElevation)))
           (if (and elev-min (= (type elev-min) 'VARIANT))
             (setq elev-min (vlax-variant-value elev-min)))

           (setq h-raw (hgl:try-prop vla '(HorizontalScale GraphScale)))
           (if (and h-raw (= (type h-raw) 'VARIANT)) (setq h-raw (vlax-variant-value h-raw)))

           (setq v-raw (hgl:try-prop vla '(VerticalScale VerticalExaggeration)))
           (if (and v-raw (= (type v-raw) 'VARIANT)) (setq v-raw (vlax-variant-value v-raw)))

           (setq rl-raw (hgl:try-prop vla '(SwapedViewDirection IsReversed IsFlipped)))
           (if (and rl-raw (= (type rl-raw) 'VARIANT)) (setq rl-raw (vlax-variant-value rl-raw)))
           (setq is-rl (cond
             ((= rl-raw :vlax-true) T)
             ((eq rl-raw T) T)
             ((and rl-raw (hgl:num-p rl-raw) (/= rl-raw 0)) T)
             (T nil)))

           (if (and ox oy (hgl:num-p sta-start) (hgl:num-p elev-min)
                    (and h-raw (hgl:num-p h-raw) (> h-raw 0))
                    (and v-raw (hgl:num-p v-raw) (> v-raw 0)))
             (list ox oy (float sta-start) (float elev-min)
                   (if is-rl (- (/ 1.0 h-raw)) (/ 1.0 h-raw))
                   (/ 1.0 v-raw) is-rl)
             nil))))
      (if (vl-catch-all-error-p res)
        (progn (princ (strcat "\n  COM read error: " (vl-catch-all-error-message res))) nil)
        res))))

;; -------------------------------------------------------------------
;; COORDINATE TRANSFORM
;; -------------------------------------------------------------------

(defun hgl:sta->x (sta sta-datum ox h-scale)
  (+ ox (* (- sta sta-datum) h-scale)))

(defun hgl:elev->y (elev elev-datum oy v-scale)
  (+ oy (* (- elev elev-datum) v-scale)))

;; -------------------------------------------------------------------
;; GETREAL WITH DEFAULT  (returns default on Enter, nil on Escape)
;; -------------------------------------------------------------------

(defun hgl:prompt-real (msg default / v)
  (setq v (getreal (strcat "\n" msg " <" (rtos default 2 4) ">: ")))
  (if v v default))

;; -------------------------------------------------------------------
;; MAIN DRAW COMMAND
;; -------------------------------------------------------------------

(defun c:HGLDRAW (/ xlsx data anno-scale pv-names pv-name pv-rows pv-ent pv-params
                    cur-ox cur-oy cur-sta-datum cur-elev-datum
                    cur-h-scale cur-v-scale cur-h-denom cur-v-denom cur-is-rl
                    prev-ox prev-oy prev-sta-datum prev-elev-datum
                    prev-h-denom prev-v-denom prev-dir
                    pts drow pipe ds-sta us-sta ds-design ds us us-design
                    lo-sta lo-hgl hi-sta hi-hgl lx ly hx hy cur-pt
                    stub-cx stub-cy stub-half
                    echo-save origin dir-str)

  (vl-load-com)

  ;; 1. Excel file --------------------------------------------------
  (setq xlsx (hgl:get-excel-path))
  (if (not (and xlsx (findfile xlsx)))
    (progn
      (setq xlsx (getfiled "Select HGL Excel File" "" "xlsx;xls" 0))
      (if xlsx (hgl:set-excel-path xlsx))))
  (if (not xlsx) (progn (princ "\nCancelled.") (exit)))

  ;; 2. Read Excel --------------------------------------------------
  (princ (strcat "\nReading: " xlsx))
  (setq data (hgl:read-excel xlsx))
  (if (not data) (progn (princ "\nERROR: No valid HGL data found in Excel.") (exit)))
  (setq pv-names (hgl:unique-pv-names data))
  (princ (strcat "\nRead " (itoa (length data)) " pipe row(s) across "
                 (itoa (length pv-names)) " profile view(s):"))
  (foreach pv-name pv-names (princ (strcat "\n  " pv-name)))

  ;; 3. Drawing scale (one value applies to the whole drawing) -------
  (setq anno-scale (hgl:prompt-real
    "Civil 3D drawing scale denominator (e.g. 20 for 1:20)"
    (max 1.0 (getvar "CANNOSCALEVALUE"))))
  (if (not anno-scale) (progn (princ "\nCancelled.") (exit)))

  ;; 4. Carry-over defaults (populated after each profile view) ------
  (setq prev-ox nil prev-oy nil prev-sta-datum nil prev-elev-datum nil
        prev-h-denom 50.0 prev-v-denom 10.0 prev-dir "L")

  ;; Suppress PLINE command echo for all drawing calls
  (setq echo-save (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)

  ;; 5. Loop over each profile view ---------------------------------
  (foreach pv-name pv-names

    ;; Rows for this profile view
    (setq pv-rows (vl-remove-if-not
      '(lambda (r) (= (hgl:trim (nth 2 r)) pv-name))
      data))
    (princ (strcat "\n\n--- Profile View: \"" pv-name "\" ("
                   (itoa (length pv-rows)) " pipe(s)) ---"))

    ;; Try auto-read by name
    (setq pv-ent (hgl:find-pv-by-name pv-name))
    (setq pv-params nil cur-h-scale nil cur-v-scale nil cur-is-rl nil)
    (setq cur-ox prev-ox cur-oy prev-oy
          cur-sta-datum prev-sta-datum cur-elev-datum prev-elev-datum)

    (if pv-ent
      (progn
        (setq pv-params (hgl:pv-read pv-ent))
        (if pv-params
          (progn
            (setq cur-ox        (nth 0 pv-params)
                  cur-oy        (nth 1 pv-params)
                  cur-sta-datum (nth 2 pv-params)
                  cur-elev-datum (nth 3 pv-params)
                  cur-h-scale   (nth 4 pv-params)
                  cur-v-scale   (nth 5 pv-params)
                  cur-is-rl     (nth 6 pv-params))
            (princ "\n  Auto-read OK:")
            (princ (strcat "\n    Origin:    (" (rtos cur-ox 2 2) ", " (rtos cur-oy 2 2) ")"))
            (princ (strcat "\n    Sta datum: " (rtos cur-sta-datum 2 2)))
            (princ (strcat "\n    Elev datum:" (rtos cur-elev-datum 2 2)))
            (princ (strcat "\n    H-scale:   1\"=" (rtos (/ anno-scale (abs cur-h-scale)) 2 1) "'"))
            (princ (strcat "\n    Direction: " (if cur-is-rl "R-to-L" "L-to-R"))))
          (princ (strcat "\n  Found but COM read failed - enter manually."))))
      (princ (strcat "\n  \"" pv-name "\" not found in drawing - enter manually.")))

    ;; Confirm / override: origin
    (if cur-ox
      (progn
        (setq origin (getpoint
          (strcat "\nOrigin [" (rtos cur-ox 2 2) "," (rtos cur-oy 2 2)
                  "] (Enter=keep, pick=override): ")))
        (if origin (setq cur-ox (car origin) cur-oy (cadr origin))))
      (progn
        (setq origin (getpoint "\nPick bottom-left origin of profile view: "))
        (if (not origin)
          (progn (setvar "CMDECHO" echo-save) (princ "\nCancelled.") (exit)))
        (setq cur-ox (car origin) cur-oy (cadr origin))))

    ;; Datum station
    (if (not cur-sta-datum) (setq cur-sta-datum (if prev-sta-datum prev-sta-datum 0.0)))
    (setq cur-sta-datum (hgl:prompt-real "Datum station (left edge of profile)" cur-sta-datum))
    (if (not cur-sta-datum)
      (progn (setvar "CMDECHO" echo-save) (princ "\nCancelled.") (exit)))

    ;; Datum elevation
    (if (not cur-elev-datum) (setq cur-elev-datum (if prev-elev-datum prev-elev-datum 0.0)))
    (setq cur-elev-datum (hgl:prompt-real "Datum elevation (bottom of profile view)" cur-elev-datum))
    (if (not cur-elev-datum)
      (progn (setvar "CMDECHO" echo-save) (princ "\nCancelled.") (exit)))

    ;; H-scale denominator
    (setq cur-h-denom
      (if cur-h-scale (/ anno-scale (abs cur-h-scale)) prev-h-denom))
    (setq cur-h-denom (hgl:prompt-real
      (strcat "H-scale denominator (e.g. 50 = 1\"=50'; scale 1:"
              (rtos anno-scale 2 0) " applied)")
      cur-h-denom))
    (if (not cur-h-denom)
      (progn (setvar "CMDECHO" echo-save) (princ "\nCancelled.") (exit)))

    ;; V-scale denominator
    (setq cur-v-denom
      (if cur-v-scale (/ anno-scale cur-v-scale) prev-v-denom))
    (setq cur-v-denom (hgl:prompt-real
      (strcat "V-scale denominator (e.g. 10 = 1\"=10'; scale 1:"
              (rtos anno-scale 2 0) " applied)")
      cur-v-denom))
    (if (not cur-v-denom)
      (progn (setvar "CMDECHO" echo-save) (princ "\nCancelled.") (exit)))

    ;; Direction
    (setq dir-str (if cur-is-rl "R" prev-dir))
    (setq dir-str (hgl:trim (getstring
      (strcat "\nDirection (L=left-to-right, R=right-to-left) <" dir-str ">: "))))
    (if (= dir-str "") (setq dir-str (if cur-is-rl "R" prev-dir)))
    (setq dir-str (strcase dir-str))

    ;; Final model-space scales
    (setq cur-h-scale (/ anno-scale cur-h-denom))
    (setq cur-v-scale (/ anno-scale cur-v-denom))
    (if (= dir-str "R") (setq cur-h-scale (- cur-h-scale)))

    ;; Update carry-over defaults for the next profile view
    (setq prev-ox cur-ox         prev-oy cur-oy
          prev-sta-datum cur-sta-datum  prev-elev-datum cur-elev-datum
          prev-h-denom cur-h-denom      prev-v-denom cur-v-denom
          prev-dir dir-str)

    ;; Sort rows: L-R ascending by min station; R-L descending by max station
    (setq pv-rows (vl-sort pv-rows
      (if (< cur-h-scale 0)
        '(lambda (a b) (> (max (nth 3 a) (nth 4 a)) (max (nth 3 b) (nth 4 b))))
        '(lambda (a b) (< (min (nth 3 a) (nth 4 a)) (min (nth 3 b) (nth 4 b)))))))

    ;; Build HGL point list
    (setq pts '())
    (foreach drow pv-rows
      (setq ds-sta (nth 3 drow)  us-sta (nth 4 drow)
            ds     (nth 6 drow)  us     (nth 7 drow))

      (if (<= ds-sta us-sta)
        (setq lo-sta ds-sta  lo-hgl ds  hi-sta us-sta  hi-hgl us)
        (setq lo-sta us-sta  lo-hgl us  hi-sta ds-sta  hi-hgl ds))

      (setq lx (hgl:sta->x  lo-sta cur-sta-datum cur-ox cur-h-scale)
            ly (hgl:elev->y lo-hgl cur-elev-datum cur-oy cur-v-scale)
            hx (hgl:sta->x  hi-sta cur-sta-datum cur-ox cur-h-scale)
            hy (hgl:elev->y hi-hgl cur-elev-datum cur-oy cur-v-scale))

      (if (< cur-h-scale 0)
        (progn  ;; R-L: hi-station end is leftmost in drawing
          (setq cur-pt (list hx hy))
          (if (or (null pts) (not (equal (last pts) cur-pt 1e-6)))
            (setq pts (append pts (list cur-pt))))
          (setq pts (append pts (list (list lx ly)))))
        (progn  ;; L-R: lo-station end is leftmost
          (setq cur-pt (list lx ly))
          (if (or (null pts) (not (equal (last pts) cur-pt 1e-6)))
            (setq pts (append pts (list cur-pt))))
          (setq pts (append pts (list (list hx hy)))))))

    ;; Draw HGL polyline
    (if (>= (length pts) 2)
      (progn
        (princ (strcat "\n  Drawing HGL polyline: " (itoa (length pts)) " vertices"))
        (command "._PLINE")
        (foreach p pts (command p))
        (command ""))
      (princ "\n  WARNING: fewer than 2 points computed - skipping polyline."))

    ;; Draw Design HGL stubs (3-ft horizontal line centered on the station)
    (setq stub-half (* 1.5 (abs cur-h-scale)))
    (foreach drow pv-rows
      (setq ds-sta    (nth 3 drow)
            us-sta    (nth 4 drow)
            ds-design (nth 5 drow)
            us-design (nth 8 drow))

      (if ds-design
        (progn
          (setq stub-cx (hgl:sta->x  ds-sta    cur-sta-datum  cur-ox cur-h-scale)
                stub-cy (hgl:elev->y ds-design  cur-elev-datum cur-oy cur-v-scale))
          (command "._PLINE"
            (list (- stub-cx stub-half) stub-cy)
            (list (+ stub-cx stub-half) stub-cy) "")))

      (if us-design
        (progn
          (setq stub-cx (hgl:sta->x  us-sta    cur-sta-datum  cur-ox cur-h-scale)
                stub-cy (hgl:elev->y us-design  cur-elev-datum cur-oy cur-v-scale))
          (command "._PLINE"
            (list (- stub-cx stub-half) stub-cy)
            (list (+ stub-cx stub-half) stub-cy) ""))))

  ) ;; end foreach pv-name

  (setvar "CMDECHO" echo-save)
  (princ "\nHGLDRAW complete.")
  (princ))

;; -------------------------------------------------------------------
;; SET EXCEL PATH COMMAND
;; -------------------------------------------------------------------

(defun c:HGLSET (/ p cur)
  (vl-load-com)
  (setq cur (hgl:get-excel-path))
  (setq p (getfiled "Select HGL Excel File" (if cur cur "") "xlsx;xls" 0))
  (if p
    (progn (hgl:set-excel-path p)
           (princ (strcat "\nHGL Excel path saved: " p)))
    (princ "\nCancelled."))
  (princ))

(princ "\n+-------------------------------------------+")
(princ "\n|  HGL Draw Routine Loaded                  |")
(princ "\n|  HGLSET  - Set Excel file path            |")
(princ "\n|  HGLDRAW - Draw all HGL polylines         |")
(princ "\n+-------------------------------------------+")
(princ)
