;; =====================================================================
;; HGL Profile Polyline - halff_hgl_draw.lsp
;;
;; Excel columns: A=Pipe, B=Start Station (DS end), C=End Station (US end),
;;                D=DS HGL, E=US HGL.  Row 1 = header; data starts row 2.
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
;; Returns list of (pipe start-sta end-sta ds-hgl us-hgl).
;; Stops on the first row with an empty Start Station cell.
;; -------------------------------------------------------------------

(defun hgl:read-excel (xlsx / pair xl wb ws nrows row data
                             pipe sta-s sta-e ds us)
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
        (setq sta-s (hgl:cell-num ws row 2))
        (if (not sta-s)
          (setq row (1+ nrows))
          (progn
            (setq pipe  (hgl:cell-str ws row 1)
                  sta-e (hgl:cell-num ws row 3)
                  ds    (hgl:cell-num ws row 4)
                  us    (hgl:cell-num ws row 5))
            (if (and sta-e ds us)
              (setq data (append data (list (list pipe sta-s sta-e ds us)))))
            (setq row (1+ row)))))
      (hgl:xl-close xl wb)
      data)))

;; -------------------------------------------------------------------
;; CIVIL 3D PROFILE VIEW AUTO-READ
;;
;; Reads origin, datum station, datum elevation, and scales from an
;; AECC_PROFILE_VIEW entity.  Civil 3D HorizontalScale / VerticalScale
;; are expressed as real-units-per-drawing-unit (e.g. 50 for 1"=50'),
;; so we invert them to produce drawing-units-per-real-unit for the
;; coordinate transform.
;;
;; Multiple property name candidates are tried in order to accommodate
;; different Civil 3D API versions.
;;
;; Returns (origin-x origin-y sta-datum elev-datum h-scale v-scale) or nil.
;; -------------------------------------------------------------------

(defun hgl:pv-read (ent / vla loc ox oy sta-start elev-min h-raw v-raw res)
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

           (if (and ox oy
                    (hgl:num-p sta-start)
                    (hgl:num-p elev-min)
                    (and h-raw (hgl:num-p h-raw) (> h-raw 0))
                    (and v-raw (hgl:num-p v-raw) (> v-raw 0)))
             (list ox oy
                   (float sta-start) (float elev-min)
                   (/ 1.0 h-raw)     (/ 1.0 v-raw))
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

(defun c:HGLDRAW (/ xlsx data ent pv-data
                    ox oy sta-datum elev-datum h-scale v-scale layer
                    pts row pipe sta-s sta-e ds us
                    lo-sta lo-hgl hi-sta hi-hgl lx ly hx hy
                    cur-pt echo-save ent-hgl origin)

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

  ;; 3. Civil 3D profile view auto-read (optional) ------------------
  (princ "\n--- Profile View Parameters ---")
  (princ "\nSelect Civil 3D Profile View to auto-read parameters")
  (princ "\n  (press Enter or Esc to skip and enter values manually): ")
  (setq ent (car (entsel "")))
  (setq pv-data nil ox nil oy nil sta-datum nil elev-datum nil
        h-scale nil v-scale nil)
  (if ent
    (progn
      (setq pv-data (hgl:pv-read ent))
      (if pv-data
        (progn
          (setq ox         (nth 0 pv-data)
                oy         (nth 1 pv-data)
                sta-datum  (nth 2 pv-data)
                elev-datum (nth 3 pv-data)
                h-scale    (nth 4 pv-data)
                v-scale    (nth 5 pv-data))
          (princ "\n  Auto-read from profile view:")
          (princ (strcat "\n    Origin:        (" (rtos ox 2 4) ", " (rtos oy 2 4) ")"))
          (princ (strcat "\n    Datum station: " (rtos sta-datum 2 4)))
          (princ (strcat "\n    Datum elev:    " (rtos elev-datum 2 4)))
          (princ (strcat "\n    H-scale:       1/" (rtos (/ 1.0 h-scale) 2 1)
                         " (drawing units per station unit)"))
          (princ (strcat "\n    V-scale:       1/" (rtos (/ 1.0 v-scale) 2 1)
                         " (drawing units per elevation unit)"))
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

  (if (not h-scale) (setq h-scale 0.02))
  (setq h-scale (hgl:prompt-real
    "H-scale (drawing units per station unit, e.g. 1/50=0.02 for 1\"=50')" h-scale))
  (if (not h-scale) (progn (princ "\nCancelled.") (exit)))

  (if (not v-scale) (setq v-scale 0.1))
  (setq v-scale (hgl:prompt-real
    "V-scale (drawing units per elevation unit, e.g. 1/10=0.1 for 1\"=10')" v-scale))
  (if (not v-scale) (progn (princ "\nCancelled.") (exit)))

  ;; 5. Layer -------------------------------------------------------
  (setq layer (getstring "\nLayer name for HGL polyline <HGL>: "))
  (if (or (not layer) (= (hgl:trim layer) "")) (setq layer "HGL"))
  (hgl:ensure-layer layer)

  ;; 6. Build point list --------------------------------------------
  ;; Sort by the lower of the two station values so the polyline runs
  ;; monotonically from the lowest station to the highest regardless of
  ;; which direction pipes are listed in the spreadsheet.
  (setq data (vl-sort data
    '(lambda (a b)
       (< (min (nth 1 a) (nth 2 a))
          (min (nth 1 b) (nth 2 b))))))

  (setq pts '())
  (foreach row data
    (setq sta-s (nth 1 row)
          sta-e (nth 2 row)
          ds    (nth 3 row)
          us    (nth 4 row))

    ;; Orient so lo-* is the downstream (lower-station) end.
    ;; The spreadsheet labels col B as DS end and col C as US end, so:
    ;;   normal  (sta-s <= sta-e): lo = DS(sta-s, ds), hi = US(sta-e, us)
    ;;   reversed (sta-s > sta-e): lo = US(sta-e, us), hi = DS(sta-s, ds)
    (if (<= sta-s sta-e)
      (setq lo-sta sta-s  lo-hgl ds  hi-sta sta-e  hi-hgl us)
      (setq lo-sta sta-e  lo-hgl us  hi-sta sta-s  hi-hgl ds))

    (setq lx (hgl:sta->x  lo-sta sta-datum ox h-scale)
          ly (hgl:elev->y lo-hgl elev-datum oy v-scale)
          hx (hgl:sta->x  hi-sta sta-datum ox h-scale)
          hy (hgl:elev->y hi-hgl elev-datum oy v-scale))

    ;; Skip the lo-station point if it matches the last point already
    ;; added (shared node between adjacent pipes).
    (setq cur-pt (list lx ly))
    (if (or (null pts) (not (equal (last pts) cur-pt 1e-6)))
      (setq pts (append pts (list cur-pt))))

    (setq pts (append pts (list (list hx hy)))))

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
