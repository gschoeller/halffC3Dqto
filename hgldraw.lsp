;; =====================================================================
;; HGL Profile Polyline - hgldraw.lsp
;;
;; Excel columns (row 1 = header, data starts row 2):
;;   A: Pipe
;;   B: Alignment
;;   C: Profile View Name
;;   D: DS Station
;;   E: US Station
;;   F: DS Design HGL  (optional; hydraulic jump at DS endpoint + 3-ft extension)
;;   G: DS HGL
;;   H: US HGL
;;   I: US Design HGL  (optional; hydraulic jump at US endpoint + 3-ft extension)
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
;; Design HGL values are hydraulic jumps at the endpoints of the
;; polyline and are woven into the main polyline as vertical segments
;; with a 3-ft horizontal extension outward:
;;   DS Design HGL -> outer stub at (min-sta - 1.5), junction at min-sta,
;;                   then vertical jump up/down to first HGL vertex
;;   US Design HGL -> vertical jump from last HGL vertex, junction at max-sta,
;;                   then outer stub at (max-sta + 1.5)
;; If a Design HGL appears at any station that is not the overall
;; minimum (DS) or maximum (US) station of the profile view group,
;; HGLDRAW reports an error and skips that profile view.
;;
;; Commands:
;;   HGLSET    - browse to and save the Excel file path
;;   HGLDRAW   - read Excel, draw all HGL polylines
;;   HGLPVTEST - diagnostic: dump COM properties of a selected profile view
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
;; GetBoundingBox returns physical extents via SAFEARRAY by-ref params.
;; hgl:bbox-var is a top-level helper so vlax-make-safearray is never
;; called from inside a quoted lambda (avoids VLisp compiler issues).
;;
;; Returns (ox oy sta-datum elev-datum h-mag v-mag nil max-ox) or nil.
;; hgl:pv-find-xy kept for HGLPVTEST compatibility.
;; -------------------------------------------------------------------

(defun hgl:pv-find-xy (vla station elevation / xv yv okv res xval yval)
  (setq xv  (vlax-make-variant 0.0 vlax-vbDouble)
        yv  (vlax-make-variant 0.0 vlax-vbDouble)
        okv (vlax-make-variant :vlax-false vlax-vbBoolean))
  (setq res (vl-catch-all-apply
    'vlax-invoke-method
    (list vla 'FindXYAtStationAndElevation
          (float station) (float elevation) xv yv okv)))
  (if (vl-catch-all-error-p res)
    nil
    (progn
      (setq xval (vlax-variant-value xv)
            yval (vlax-variant-value yv))
      (if (and (hgl:num-p xval) (hgl:num-p yval))
        (list (float xval) (float yval))
        nil))))

(defun hgl:bbox-var ()
  (vlax-make-variant (vlax-make-safearray vlax-vbDouble (cons 0 2))))

(defun hgl:pv-read (ent / vla sta-s sta-e elv-n elv-x
                          bb-lo bb-hi lo-x lo-y hi-x hi-y hm vm res)
  (if (/= (cdr (assoc 0 (entget ent))) "AECC_PROFILE_VIEW")
    (progn (princ "\n  Not an AECC_PROFILE_VIEW.") nil)
    (progn
      (setq bb-lo (hgl:bbox-var))
      (setq bb-hi (hgl:bbox-var))
      (setq res (vl-catch-all-apply
        '(lambda ()
           (setq vla (vlax-ename->vla-object ent))
           (setq sta-s (hgl:variant->val (vlax-get-property vla 'StationStart)))
           (setq sta-e (hgl:variant->val (vlax-get-property vla 'StationEnd)))
           (setq elv-n (hgl:variant->val (vlax-get-property vla 'ElevationMin)))
           (setq elv-x (hgl:variant->val (vlax-get-property vla 'ElevationMax)))
           (if (not (and (hgl:num-p sta-s) (hgl:num-p sta-e)
                         (hgl:num-p elv-n) (hgl:num-p elv-x)
                         (/= sta-s sta-e) (/= elv-n elv-x)))
             nil
             (progn
               (vlax-invoke-method vla 'GetBoundingBox bb-lo bb-hi)
               (setq lo-x (car  (vlax-safearray->list (vlax-variant-value bb-lo)))
                     lo-y (cadr (vlax-safearray->list (vlax-variant-value bb-lo)))
                     hi-x (car  (vlax-safearray->list (vlax-variant-value bb-hi)))
                     hi-y (cadr (vlax-safearray->list (vlax-variant-value bb-hi))))
               (setq hm (/ (- hi-x lo-x) (abs (- sta-e sta-s))))
               (setq vm (/ (- hi-y lo-y) (abs (- elv-x elv-n))))
               (if (and (> hm 0) (> vm 0))
                 (list (float lo-x) (float lo-y)
                       (float sta-s) (float elv-n)
                       (float hm) (float vm) nil
                       (float hi-x))
                 nil)))))
      (if (vl-catch-all-error-p res)
        (progn (princ (strcat "\n  COM read error: " (vl-catch-all-error-message res))) nil)
        res)))))

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
                    cur-h-scale cur-v-scale cur-h-denom cur-v-denom cur-is-rl cur-align
                    prev-ox prev-oy prev-sta-datum prev-elev-datum
                    prev-h-denom prev-v-denom prev-dir
                    pts drow pipe ds-sta us-sta ds-design ds us us-design
                    lo-sta lo-hgl hi-sta hi-hgl lx ly hx hy cur-pt
                    min-sta max-sta ds-design-val us-design-val design-err
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

    ;; Alignment name (col B) from first row of this group
    (setq cur-align (hgl:trim (nth 1 (car pv-rows))))

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
          (princ (strcat "\n  Alignment \"" cur-align
                         "\": found profile view but COM read failed - enter manually."))))
      (princ (strcat "\n  Alignment \"" cur-align
                     "\": profile view not found in drawing - enter manually.")))

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

    ;; Compute overall station range for this profile view group.
    (setq min-sta (apply 'min (mapcar '(lambda (r) (min (nth 3 r) (nth 4 r))) pv-rows)))
    (setq max-sta (apply 'max (mapcar '(lambda (r) (max (nth 3 r) (nth 4 r))) pv-rows)))

    ;; Validate Design HGL positions.
    (setq design-err nil ds-design-val nil us-design-val nil)
    (foreach drow pv-rows
      (setq ds-design (nth 5 drow)
            us-design (nth 8 drow))
      (if ds-design
        (if (equal (nth 3 drow) min-sta 1e-4)
          (setq ds-design-val ds-design)
          (progn
            (princ (strcat "\nERROR: DS Design HGL on pipe " (nth 0 drow)
                           " (DS station " (rtos (nth 3 drow) 2 2)
                           ") is not at the start of the line (min station "
                           (rtos min-sta 2 2) "). Skipping \"" pv-name "\"."))
            (setq design-err T))))
      (if us-design
        (if (equal (nth 4 drow) max-sta 1e-4)
          (setq us-design-val us-design)
          (progn
            (princ (strcat "\nERROR: US Design HGL on pipe " (nth 0 drow)
                           " (US station " (rtos (nth 4 drow) 2 2)
                           ") is not at the end of the line (max station "
                           (rtos max-sta 2 2) "). Skipping \"" pv-name "\"."))
            (setq design-err T)))))

    (if (not design-err)
      (progn
        ;; Sort rows: L-R ascending by min station; R-L descending by max station
        (setq pv-rows (vl-sort pv-rows
          (if (< cur-h-scale 0)
            '(lambda (a b) (> (max (nth 3 a) (nth 4 a)) (max (nth 3 b) (nth 4 b))))
            '(lambda (a b) (< (min (nth 3 a) (nth 4 a)) (min (nth 3 b) (nth 4 b)))))))

        ;; Build main HGL point list.
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
            (progn  ;; R-L: hi-station end is leftmost
              (setq cur-pt (list hx hy))
              (if (or (null pts) (not (equal (last pts) cur-pt 1e-6)))
                (setq pts (append pts (list cur-pt))))
              (setq pts (append pts (list (list lx ly)))))
            (progn  ;; L-R: lo-station end is leftmost
              (setq cur-pt (list lx ly))
              (if (or (null pts) (not (equal (last pts) cur-pt 1e-6)))
                (setq pts (append pts (list cur-pt))))
              (setq pts (append pts (list (list hx hy)))))))

        ;; Weave Design HGL jumps into the polyline endpoints, each with a
        ;; 3-ft (1.5 station-unit) horizontal outer extension.
        ;;
        ;; L-R layout:
        ;;   [outer@min-1.5, junction@min, ..run.., junction@max, outer@max+1.5]
        ;; R-L layout (negative h-scale flips X automatically):
        ;;   [outer@max+1.5, junction@max, ..run.., junction@min, outer@min-1.5]
        (if (< cur-h-scale 0)
          (progn
            ;; R-L beginning = US end (max-sta is leftmost)
            (if us-design-val
              (progn
                (setq pts (cons
                  (list (hgl:sta->x max-sta cur-sta-datum cur-ox cur-h-scale)
                        (hgl:elev->y us-design-val cur-elev-datum cur-oy cur-v-scale))
                  pts))
                (setq pts (cons
                  (list (hgl:sta->x (+ max-sta 1.5) cur-sta-datum cur-ox cur-h-scale)
                        (hgl:elev->y us-design-val cur-elev-datum cur-oy cur-v-scale))
                  pts))))
            ;; R-L end = DS end (min-sta is rightmost)
            (if ds-design-val
              (progn
                (setq pts (append pts
                  (list (list (hgl:sta->x min-sta cur-sta-datum cur-ox cur-h-scale)
                              (hgl:elev->y ds-design-val cur-elev-datum cur-oy cur-v-scale)))))
                (setq pts (append pts
                  (list (list (hgl:sta->x (- min-sta 1.5) cur-sta-datum cur-ox cur-h-scale)
                              (hgl:elev->y ds-design-val cur-elev-datum cur-oy cur-v-scale))))))))
          (progn
            ;; L-R beginning = DS end (min-sta is leftmost)
            (if ds-design-val
              (progn
                (setq pts (cons
                  (list (hgl:sta->x min-sta cur-sta-datum cur-ox cur-h-scale)
                        (hgl:elev->y ds-design-val cur-elev-datum cur-oy cur-v-scale))
                  pts))
                (setq pts (cons
                  (list (hgl:sta->x (- min-sta 1.5) cur-sta-datum cur-ox cur-h-scale)
                        (hgl:elev->y ds-design-val cur-elev-datum cur-oy cur-v-scale))
                  pts))))
            ;; L-R end = US end (max-sta is rightmost)
            (if us-design-val
              (progn
                (setq pts (append pts
                  (list (list (hgl:sta->x max-sta cur-sta-datum cur-ox cur-h-scale)
                              (hgl:elev->y us-design-val cur-elev-datum cur-oy cur-v-scale)))))
                (setq pts (append pts
                  (list (list (hgl:sta->x (+ max-sta 1.5) cur-sta-datum cur-ox cur-h-scale)
                              (hgl:elev->y us-design-val cur-elev-datum cur-oy cur-v-scale)))))))))

        ;; Draw the complete polyline
        (if (>= (length pts) 2)
          (progn
            (princ (strcat "\n  Drawing HGL polyline: " (itoa (length pts)) " vertices"))
            (command "._PLINE")
            (foreach p pts (command p))
            (command ""))
          (princ "\n  WARNING: fewer than 2 points - skipping polyline."))))

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

;; -------------------------------------------------------------------
;; HGLPVTEST helpers - top-level so strcat/symbol issues are avoided
;; -------------------------------------------------------------------

(defun hgl:probe-prop (obj pname / r pstr)
  ;; pname may be a symbol or string; convert for display
  (setq pstr (if (= (type pname) 'STR) pname (vl-princ-to-string pname)))
  (setq r (vl-catch-all-apply 'vlax-get-property (list obj pname)))
  (if (vl-catch-all-error-p r)
    (progn
      (princ (strcat "\n    " pstr ": ERROR - " (vl-catch-all-error-message r)))
      nil)
    (progn
      (princ (strcat "\n    " pstr ": " (vl-princ-to-string r)))
      r)))

;; -------------------------------------------------------------------
;; HGLPVTEST - diagnostic command to debug COM read failures
;;
;; Select any AECC_PROFILE_VIEW entity and this command will:
;;   1. Confirm the DXF entity type
;;   2. Probe every property name that hgl:pv-read uses and report
;;      the raw value returned (or the error message on failure)
;;   3. Call vlax-dump-object for a full COM property/method listing
;; -------------------------------------------------------------------

(defun c:HGLPVTEST (/ ent etype vla
                      pv-sta-s pv-sta-e pv-elv-n pv-elv-x
                      pv-xy1 pv-xy2 pv-xy3 pv-hsc pv-vsc)

  (vl-load-com)

  (princ "\nHGLPVTEST: Select a profile view entity...")
  (setq ent (car (entsel "\nSelect profile view: ")))
  (if (not ent)
    (progn (princ "\nCancelled.") (exit)))

  ;; DXF type
  (setq etype (cdr (assoc 0 (entget ent))))
  (princ (strcat "\n\nEntity DXF type: " etype))
  (if (/= etype "AECC_PROFILE_VIEW")
    (princ "\n  WARNING: Not an AECC_PROFILE_VIEW - results may be incomplete."))

  ;; Get VLA object
  (setq vla (vl-catch-all-apply 'vlax-ename->vla-object (list ent)))
  (if (vl-catch-all-error-p vla)
    (progn
      (princ (strcat "\n  FATAL: vlax-ename->vla-object failed: "
                     (vl-catch-all-error-message vla)))
      (exit)))
  (princ "\n  VLA object obtained OK.")

  ;; Probe individual properties
  (princ "\n\n--- Probing COM properties ---")

  (princ "\n  [Name]")
  (hgl:probe-prop vla 'Name)

  (princ "\n  [Stations]")
  (hgl:probe-prop vla 'StationStart)
  (hgl:probe-prop vla 'StationEnd)

  (princ "\n  [Elevations]")
  (hgl:probe-prop vla 'ElevationMin)
  (hgl:probe-prop vla 'ElevationMax)

  (princ "\n  [Scale properties - for reference]")
  (hgl:probe-prop vla 'VerticalScale)
  (hgl:probe-prop vla 'HorizontalScale)

  (princ "\n  [Location properties - expected to fail in v1.2+]")
  (hgl:probe-prop vla 'Location)
  (hgl:probe-prop vla 'InsertionPoint)
  (hgl:probe-prop vla 'Origin)

  ;; Test FindXYAtStationAndElevation with the known station/elevation values
  (princ "\n\n--- Testing FindXYAtStationAndElevation ---")
  (setq pv-sta-s (hgl:variant->val (vl-catch-all-apply 'vlax-get-property (list vla 'StationStart)))
        pv-sta-e (hgl:variant->val (vl-catch-all-apply 'vlax-get-property (list vla 'StationEnd)))
        pv-elv-n (hgl:variant->val (vl-catch-all-apply 'vlax-get-property (list vla 'ElevationMin)))
        pv-elv-x (hgl:variant->val (vl-catch-all-apply 'vlax-get-property (list vla 'ElevationMax))))
  (if (and (hgl:num-p pv-sta-s) (hgl:num-p pv-sta-e)
           (hgl:num-p pv-elv-n) (hgl:num-p pv-elv-x))
    (progn
      (princ (strcat "\n  StaStart=" (rtos pv-sta-s 2 2)
                     "  StaEnd=" (rtos pv-sta-e 2 2)
                     "  ElevMin=" (rtos pv-elv-n 2 2)
                     "  ElevMax=" (rtos pv-elv-x 2 2)))
      (setq pv-xy1 (hgl:pv-find-xy vla pv-sta-s pv-elv-n)
            pv-xy2 (hgl:pv-find-xy vla pv-sta-e pv-elv-n)
            pv-xy3 (hgl:pv-find-xy vla pv-sta-s pv-elv-x))
      (if pv-xy1
        (princ (strcat "\n  FindXY(StaStart,ElevMin) = ("
                       (rtos (car pv-xy1) 2 4) ", " (rtos (cadr pv-xy1) 2 4) ")"))
        (princ "\n  FindXY(StaStart,ElevMin) = FAILED"))
      (if pv-xy2
        (princ (strcat "\n  FindXY(StaEnd,  ElevMin) = ("
                       (rtos (car pv-xy2) 2 4) ", " (rtos (cadr pv-xy2) 2 4) ")"))
        (princ "\n  FindXY(StaEnd,  ElevMin) = FAILED"))
      (if pv-xy3
        (princ (strcat "\n  FindXY(StaStart,ElevMax) = ("
                       (rtos (car pv-xy3) 2 4) ", " (rtos (cadr pv-xy3) 2 4) ")"))
        (princ "\n  FindXY(StaStart,ElevMax) = FAILED"))
      (if (and pv-xy1 pv-xy2 pv-xy3 (/= pv-sta-s pv-sta-e) (/= pv-elv-n pv-elv-x))
        (progn
          (setq pv-hsc (/ (- (car pv-xy2) (car pv-xy1)) (- pv-sta-e pv-sta-s)))
          (setq pv-vsc (/ (- (cadr pv-xy3) (cadr pv-xy1)) (- pv-elv-x pv-elv-n)))
          (princ (strcat "\n  => h-scale (model units/ft): " (rtos pv-hsc 2 6)))
          (princ (strcat "\n  => v-scale (model units/ft): " (rtos pv-vsc 2 6)))
          (princ (strcat "\n  => Direction: " (if (< pv-hsc 0) "R-to-L" "L-to-R")))
          (princ "\n  => Auto-read WOULD SUCCEED with these values."))
        (princ "\n  Cannot derive scales - one or more FindXY calls failed.")))
    (princ "\n  StationStart/End or ElevationMin/Max unavailable - skipping test."))

  ;; Full COM dump
  (princ "\n\n--- Full COM object dump (vlax-dump-object) ---")
  (vl-catch-all-apply 'vlax-dump-object (list vla T))

  (princ "\n\nHGLPVTEST complete.")
  (princ))

(princ "\n+-------------------------------------------+")
(princ "\n|  HGL Draw Routine Loaded                  |")
(princ "\n|  HGLSET    - Set Excel file path          |")
(princ "\n|  HGLDRAW   - Draw all HGL polylines       |")
(princ "\n|  HGLPVTEST - Debug profile view COM read  |")
(princ "\n+-------------------------------------------+")
(princ)
