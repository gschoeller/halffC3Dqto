;; =====================================================================
;; HGL Profile Polyline - halff_hgl_draw.lsp
;;
;; Excel columns: A=Pipe, B=DS Station, C=US Station,
;;                D=DS HGL, E=US HGL, F=Design Station (ignored).
;;                Row 1 = header; data starts row 2.
;;
;; Commands:
;;   HGLSET  - browse to and save the Excel file path (stored in a
;;             .hglpath sidecar file next to the DWG)
;;   HGLDRAW - read Excel, optionally auto-read a Civil 3D profile view
;;             for origin/datum/scale, prompt for any missing values,
;;             draw the HGL as a polyline
;;
;; Civil 3D integration (HGLDRAW step 3):
;;   Select an AECC_PROFILE_VIEW entity to attempt auto-read of:
;;     Location / InsertionPoint -> world XY of the bottom-left corner
;;     StationStart              -> station value at the left edge
;;     ElevationMin              -> datum elevation at the bottom
;;     HorizontalScale           -> real-units per drawing-unit (inverted)
;;     VerticalScale             -> real-units per drawing-unit (inverted)
;;   All values are displayed and can be confirmed or overridden.
;;   Press Enter at the entity prompt to skip to fully manual entry.
;;
;; Direction handling:
;;   Pipes are sorted by their lower station value before point-building,
;;   so the spreadsheet order does not matter.  For each pipe, whichever
;;   station is smaller is treated as the DS end regardless of which
;;   column it appears in.  Shared nodes between adjacent pipes are
;;   de-duplicated so the polyline has no doubled vertices.
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
;; Returns list of (pipe ds-sta us-sta ds-hgl us-hgl).
;; Col F (Design Station) is present in the sheet but not read.
;; Stops on the first row with an empty DS Station cell (col B).
;; -------------------------------------------------------------------

(defun hgl:read-excel (xlsx / pair xl wb ws nrows row data
                             pipe ds-sta us-sta ds us)
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
        (setq ds-sta (hgl:cell-num ws row 2)) ;; col B: DS Station
        (if (not ds-sta)
          (setq row (1+ nrows))
          (progn
            (setq pipe   (hgl:cell-str ws row 1)
                  us-sta (hgl:cell-num ws row 3) ;; col C: US Station
                  ds     (hgl:cell-num ws row 4) ;; col D: DS HGL
                  us     (hgl:cell-num ws row 5)) ;; col E: US HGL
                  ;; col F: Design Station - not read
            (if (and us-sta ds us)
              (setq data (append data (list (list pipe ds-sta us-sta ds us)))))
            (setq row (1+ row)))))
      (hgl:xl-close xl wb)
      data)))

;; -------------------------------------------------------------------
;; CIVIL 3D PROFILE VIEW AUTO-READ
;;
;; Reads origin, datum station, datum elevation, scales, and direction
;; from an AECC_PROFILE_VIEW entity.
;;
;; HorizontalScale / VerticalScale are real-units-per-drawing-unit
;; (e.g. 50 for 1"=50'), so we invert them.
;;
;; SwapedViewDirection (Civil 3D's own typo) is T when the profile runs
;; right-to-left.  For R-L profiles h-scale is returned negative so the
;; coordinate transform X = ox + (sta - sta-datum) * h-scale places each
;; station correctly: higher stations map to smaller X (left in drawing).
;;
;; Multiple property name candidates accommodate different API versions.
;;
;; Returns (origin-x origin-y sta-datum elev-datum h-scale v-scale is-rl)
;; where is-rl is T for right-to-left profiles, nil for left-to-right.
;; Returns nil if the entity is not a profile view or properties fail.
;; -------------------------------------------------------------------

(defun hgl:pv-read (ent / vla loc ox oy sta-start elev-min h-raw v-raw rl-raw is-rl res)
  (if (/= (cdr (assoc 0 (entget ent))) "AECC_PROFILE_VIEW")
    (progn (princ "\n  Not an AECC_PROFILE_VIEW - skipping auto-read.") nil)
    (progn
      (setq res (vl-catch-all-apply
        '(lambda ()
           (setq vla (vlax-ename->vla-object ent))

           ;; Origin - bottom-left corner in world coordinates
           (setq loc (hgl:try-prop vla '(Location InsertionPoint Origin)))
           (if (= (type loc) 'VARIANT) (setq loc (vlax-variant-value loc)))
           (cond
             ((= (type loc) 'SAFEARRAY)
              (setq ox (vlax-safearray-get-element loc 0)
                    oy (vlax-safearray-get-element loc 1)))
             ((listp loc)
              (setq ox (car loc) oy (cadr loc)))
             (T (setq ox nil oy nil)))

           ;; Station value at the left edge of the view
           (setq sta-start (hgl:try-prop vla '(StationStart StartStation)))
           (if (and sta-start (= (type sta-start) 'VARIANT))
             (setq sta-start (vlax-variant-value sta-start)))

           ;; Elevation at the bottom of the view
           (setq elev-min (hgl:try-prop vla '(ElevationMin MinimumElevation DatumElevation)))
           (if (and elev-min (= (type elev-min) 'VARIANT))
             (setq elev-min (vlax-variant-value elev-min)))

           ;; Horizontal scale (real-units per drawing-unit -> invert)
           (setq h-raw (hgl:try-prop vla '(HorizontalScale GraphScale)))
           (if (and h-raw (= (type h-raw) 'VARIANT))
             (setq h-raw (vlax-variant-value h-raw)))

           ;; Vertical scale (real-units per drawing-unit -> invert)
           (setq v-raw (hgl:try-prop vla '(VerticalScale VerticalExaggeration)))
           (if (and v-raw (= (type v-raw) 'VARIANT))
             (setq v-raw (vlax-variant-value v-raw)))

           ;; Direction: SwapedViewDirection = T means right-to-left (Civil 3D typo is intentional)
           (setq rl-raw (hgl:try-prop vla '(SwapedViewDirection IsReversed IsFlipped)))
           (if (and rl-raw (= (type rl-raw) 'VARIANT))
             (setq rl-raw (vlax-variant-value rl-raw)))
           (setq is-rl (cond
             ((= rl-raw :vlax-true) T)
             ((eq rl-raw T) T)
             ((and rl-raw (hgl:num-p rl-raw) (/= rl-raw 0)) T)
             (T nil)))

           (if (and ox oy
                    (hgl:num-p sta-start)
                    (hgl:num-p elev-min)
                    (and h-raw (hgl:num-p h-raw) (> h-raw 0))
                    (and v-raw (hgl:num-p v-raw) (> v-raw 0)))
             (list ox oy
                   (float sta-start) (float elev-min)
                   ;; Negative h-scale encodes R-L direction for the transform
                   (if is-rl (- (/ 1.0 h-raw)) (/ 1.0 h-raw))
                   (/ 1.0 v-raw)
                   is-rl)
             nil))))

      (if (vl-catch-all-error-p res)
        (progn
          (princ (strcat "\n  Profile view read error: "
                         (vl-catch-all-error-message res)))
          nil)
        res))))

;; -------------------------------------------------------------------
;; COORDINATE TRANSFORM
;; -------------------------------------------------------------------

(defun hgl:sta->x (sta sta-datum ox h-scale)
  (+ ox (* (- sta sta-datum) h-scale)))

(defun hgl:elev->y (elev elev-datum oy v-scale)
  (+ oy (* (- elev elev-datum) v-scale)))

;; -------------------------------------------------------------------
;; GETREAL WITH DEFAULT
;; Returns nil on Escape; returns default when user presses Enter.
;; -------------------------------------------------------------------

(defun hgl:prompt-real (msg default / v)
  (setq v (getreal (strcat "\n" msg " <" (rtos default 2 4) ">: ")))
  (if v v default))

;; -------------------------------------------------------------------
;; ENSURE LAYER EXISTS (green, continuous)
;; -------------------------------------------------------------------

(defun hgl:ensure-layer (lname)
  (if (not (tblsearch "LAYER" lname))
    (entmake (list '(0 . "LAYER")
                   '(100 . "AcDbSymbolTableRecord")
                   '(100 . "AcDbLayerTableRecord")
                   (cons 2 lname) '(70 . 0) '(62 . 3) '(6 . "Continuous")))))

;; -------------------------------------------------------------------
;; MAIN DRAW COMMAND
;; -------------------------------------------------------------------

(defun c:HGLDRAW (/ xlsx data ent pv-data anno-scale
                    ox oy sta-datum elev-datum h-scale v-scale h-denom v-denom is-rl layer
                    pts row pipe ds-sta us-sta ds us
                    lo-sta lo-hgl hi-sta hi-hgl lx ly hx hy
                    cur-pt echo-save ent-hgl origin dir-str)

  (vl-load-com)

  ;; 1. Excel file --------------------------------------------------
  (setq xlsx (hgl:get-excel-path))
  (if (not (and xlsx (findfile xlsx)))
    (progn
      (setq xlsx (getfiled "Select HGL Excel File" "" "xlsx;xls" 0))
      (if xlsx (hgl:set-excel-path xlsx))))
  (if (not xlsx)
    (progn (princ "\nCancelled.") (exit)))

  ;; 2. Read Excel --------------------------------------------------
  (princ (strcat "\nReading: " xlsx))
  (setq data (hgl:read-excel xlsx))
  (if (not data)
    (progn (princ "\nERROR: No valid HGL data found in Excel.") (exit)))
  (princ (strcat "\nRead " (itoa (length data)) " pipe row(s)."))

  ;; Model-space annotation scale (e.g. 20 for 1:20 drawing).
  ;; Profile view scales reported on paper must be divided by this value to
  ;; get drawing units per real unit.  The routine applies it automatically.
  (setq anno-scale (max 1.0 (getvar "CANNOSCALEVALUE")))
  (princ (strcat "\nAnnotation scale: 1:" (rtos anno-scale 2 0)))

  ;; 3. Civil 3D profile view auto-read (optional) ------------------
  (princ "\n--- Profile View Parameters ---")
  (princ "\nSelect Civil 3D Profile View to auto-read parameters")
  (princ "\n  (press Enter or Esc to skip and enter values manually): ")
  (setq ent (car (entsel "")))
  (setq pv-data nil ox nil oy nil sta-datum nil elev-datum nil
        h-scale nil v-scale nil is-rl nil)
  (if ent
    (progn
      (setq pv-data (hgl:pv-read ent))
      (if pv-data
        (progn
          (setq ox         (nth 0 pv-data)
                oy         (nth 1 pv-data)
                sta-datum  (nth 2 pv-data)
                elev-datum (nth 3 pv-data)
                h-scale    (nth 4 pv-data)  ; negative when R-L
                v-scale    (nth 5 pv-data)
                is-rl      (nth 6 pv-data))
          (princ "\n  Auto-read from profile view:")
          (princ (strcat "\n    Origin:        (" (rtos ox 2 4) ", " (rtos oy 2 4) ")"))
          (princ (strcat "\n    Datum station: " (rtos sta-datum 2 4)))
          (princ (strcat "\n    Datum elev:    " (rtos elev-datum 2 4)))
          ;; Back-calculate the paper denominator so the user sees a familiar "1\"=50'" value.
          ;; formula: denom = anno-scale / |h-scale|  (e.g. 20 / 0.4 = 50)
          (princ (strcat "\n    H-scale:       1\"=" (rtos (/ anno-scale (abs h-scale)) 2 1) "'"))
          (princ (strcat "\n    V-scale:       1\"=" (rtos (/ anno-scale v-scale) 2 1) "'"))
          (princ (strcat "\n    Direction:     "
                         (if is-rl "Right-to-Left" "Left-to-Right")))
          (princ "\n  Press Enter to accept each value or type a new one."))
        (princ "\n  Could not auto-read - enter values manually."))))

  ;; 4. Confirm / override each parameter ---------------------------

  ;; Origin
  (if ox
    (progn
      (setq origin (getpoint (strcat "\nOrigin [" (rtos ox 2 2)
                                     "," (rtos oy 2 2) "] (Enter=keep, or pick): ")))
      (if origin (setq ox (car origin) oy (cadr origin))))
    (progn
      (setq origin (getpoint "\nPick bottom-left origin of profile view: "))
      (if (not origin) (progn (princ "\nCancelled.") (exit)))
      (setq ox (car origin) oy (cadr origin))))

  (if (not sta-datum)  (setq sta-datum 0.0))
  (setq sta-datum (hgl:prompt-real
    "Datum station (station value at left edge of profile)" sta-datum))
  (if (not sta-datum) (progn (princ "\nCancelled.") (exit)))

  (if (not elev-datum) (setq elev-datum 0.0))
  (setq elev-datum (hgl:prompt-real
    "Datum elevation (elevation at bottom of profile view)" elev-datum))
  (if (not elev-datum) (progn (princ "\nCancelled.") (exit)))

  ;; H and V scale denominators.
  ;; Enter the number after "1 inch equals" from the profile view properties
  ;; (e.g. 50 for 1"=50', 100 for 1"=100').  The annotation scale
  ;; (1:{anno-scale}) is multiplied in automatically so the result lands in
  ;; model-space drawing units.  Formula: model_scale = anno_scale / denom.
  ;; If auto-read succeeded, the denominator is back-calculated from the API
  ;; value so you can verify it looks right before accepting.
  (setq h-denom (if h-scale (/ anno-scale (abs h-scale)) 50.0))
  (setq h-denom (hgl:prompt-real
    (strcat "H-scale denominator (e.g. 50 = 1\"=50'; scale 1:"
            (rtos anno-scale 2 0) " applied automatically)")
    h-denom))
  (if (not h-denom) (progn (princ "\nCancelled.") (exit)))

  (setq v-denom (if v-scale (/ anno-scale v-scale) 10.0))
  (setq v-denom (hgl:prompt-real
    (strcat "V-scale denominator (e.g. 10 = 1\"=10'; scale 1:"
            (rtos anno-scale 2 0) " applied automatically)")
    v-denom))
  (if (not v-denom) (progn (princ "\nCancelled.") (exit)))

  ;; Direction: L-to-R (normal) or R-to-L.
  ;; Default comes from auto-read flag; otherwise L-to-R.
  (setq dir-str (hgl:trim (getstring
    (strcat "\nProfile direction (L=left-to-right, R=right-to-left) <"
            (if is-rl "R" "L") ">: "))))
  (if (= dir-str "") (setq dir-str (if is-rl "R" "L")))
  (setq dir-str (strcase dir-str))

  ;; Compute final model-space scales: anno_scale / denom, signed for direction.
  (setq h-scale (/ anno-scale h-denom))
  (setq v-scale (/ anno-scale v-denom))
  (if (= dir-str "R") (setq h-scale (- h-scale)))

  ;; 5. Layer -------------------------------------------------------
  (setq layer (getstring "\nLayer name for HGL polyline <HGL>: "))
  (if (or (not layer) (= (hgl:trim layer) "")) (setq layer "HGL"))
  (hgl:ensure-layer layer)

  ;; 6. Build point list --------------------------------------------
  ;; For L-R: sort ascending by min station so the polyline builds left→right.
  ;; For R-L: sort descending by max station so it builds left→right in world
  ;;   coords (high station = left side of drawing when h-scale is negative).
  ;; Either way, the negative h-scale in the transform maps stations to the
  ;; correct world X, and shared-node deduplication compares world coords.
  (setq data (vl-sort data
    (if (< h-scale 0)
      '(lambda (a b) (> (max (nth 1 a) (nth 2 a)) (max (nth 1 b) (nth 2 b))))
      '(lambda (a b) (< (min (nth 1 a) (nth 2 a)) (min (nth 1 b) (nth 2 b)))))))

  (setq pts '())
  (foreach row data
    (setq ds-sta (nth 1 row)
          us-sta (nth 2 row)
          ds     (nth 3 row)
          us     (nth 4 row))

    ;; Split into lo-station and hi-station ends.
    ;; Col B=DS Station is normally the lower station; guard either way.
    (if (<= ds-sta us-sta)
      (setq lo-sta ds-sta  lo-hgl ds  hi-sta us-sta  hi-hgl us)
      (setq lo-sta us-sta  lo-hgl us  hi-sta ds-sta  hi-hgl ds))

    (setq lx (hgl:sta->x  lo-sta sta-datum ox h-scale)
          ly (hgl:elev->y lo-hgl elev-datum oy v-scale)
          hx (hgl:sta->x  hi-sta sta-datum ox h-scale)
          hy (hgl:elev->y hi-hgl elev-datum oy v-scale))

    ;; Append in the order that traces the polyline left-to-right in
    ;; world coordinates and keeps shared nodes contiguous.
    ;; L-R: lo-station end has smaller world X -> append lo then hi.
    ;; R-L: hi-station end has smaller world X (negative h-scale) -> append hi then lo.
    (if (< h-scale 0)
      (progn
        (setq cur-pt (list hx hy))
        (if (or (null pts) (not (equal (last pts) cur-pt 1e-6)))
          (setq pts (append pts (list cur-pt))))
        (setq pts (append pts (list (list lx ly)))))
      (progn
        (setq cur-pt (list lx ly))
        (if (or (null pts) (not (equal (last pts) cur-pt 1e-6)))
          (setq pts (append pts (list cur-pt))))
        (setq pts (append pts (list (list hx hy)))))))

  (if (< (length pts) 2)
    (progn (princ "\nERROR: Fewer than 2 points computed - check Excel data.") (exit)))

  (princ (strcat "\nDrawing HGL polyline with " (itoa (length pts))
                 " vertices on layer \"" layer "\"..."))

  ;; 7. Draw polyline -----------------------------------------------
  (setq echo-save (getvar "CMDECHO"))
  (setvar "CMDECHO" 0)
  (command "._PLINE")
  (foreach p pts (command p))
  (command "")
  (setvar "CMDECHO" echo-save)

  ;; 8. Assign layer ------------------------------------------------
  (setq ent-hgl (entlast))
  (entmod (subst (cons 8 layer) (assoc 8 (entget ent-hgl)) (entget ent-hgl)))
  (entupd ent-hgl)

  (princ (strcat "\nHGL polyline drawn on layer: " layer))
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
(princ "\n|  HGLDRAW - Draw HGL polyline from Excel   |")
(princ "\n+-------------------------------------------+")
(princ)
