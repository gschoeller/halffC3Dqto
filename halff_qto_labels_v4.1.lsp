;; =====================================================================
;; Halff QTO Labels v4.0
;;
;; Built on Halff QTO v1.22.  Adds TEXT CONTENT column support so that
;; QRUN can count TEXT, MTEXT, MLEADER, and Civil 3D label entities
;; whose text contains a specified string (case-insensitive substring).
;;
;; NEW in v1.0 Labels (vs v1.22):
;; - Excel mapping now supports a "TEXT CONTENT" column (column 9).
;;   When filled in, the row is processed in text-count mode instead of
;;   the geometry / layer mode used for all other pay items.
;; - Text-count mode iterates every entity in the target DWG, extracts
;;   text via vla-get-textstring, and counts matches that contain the
;;   search string.  Layer filter is not required for these rows.
;; - OBJECT column for text-count rows accepts:
;;     TEXT, MTEXT, MLEADER     – search those specific entity types
;;     Alignment Station Offset Label (or any unrecognised type)
;;                              – search all entities for text content
;; - VP assignment is attempted for text entities using the same
;;   center-point logic as geometry entities.  Labels that fall outside
;;   all viewports (e.g. sheet-level annotations) are counted in
;;   QTY_MODEL only.
;; - All geometry / layer / VP logic from v1.22 is unchanged.
;; NEW in v2.0:
;; - TEXT QTY column: extract and sum numbers embedded in label text.
;;   Example: TEXT CONTENT="PROP. # LF STEEL ENCASEMENT", TEXT QTY="#"
;;   Finds all matching labels, extracts the number at #, sums them.
;; =====================================================================

(vl-load-com)

;; ===========================================
;; GLOBAL VARIABLES
;; ===========================================
(setq *HALFF_VP_DEFS* nil)
(setq *HALFF_LAST_CSV* nil)
(setq *HALFF_LAST_REPORT* nil)
(setq *HALFF_ROUND_EACH* T)
(setq *HALFF_ROUND_MODE* "UP")
(setq *HALFF_FAILURE_LOG* nil)
(setq *HALFF_CROSSING_QTY_ZERO* T)
(setq *HALFF_EXCLUSION_DIST* 1.0)
(setq *HALFF_HIGHLIGHT_ENABLE* T)
(setq *HALFF_PAYITEM_ENTS* nil)
(setq *HALFF_LAST_HILITE_PAY* nil)
(setq *HALFF_LAST_HILITE_VP* nil)
(setq *HALFF_FAIL_ENTS* nil)

;; ===========================================
;; UTILITY FUNCTIONS
;; ===========================================

(defun halff:variant->list (v)
  (cond
    ((= (type v) 'VARIANT) (halff:variant->list (vlax-variant-value v)))
    ((= (type v) 'SAFEARRAY) (vlax-safearray->list v))
    (T v)))

(defun halff:nbspace->space (s)
  (if (and s (= (type s) 'STR)) (vl-string-translate (chr 160) " " s) s))

(defun halff:safe-str (x)
  (if x (if (= (type x) 'STR) x (vl-princ-to-string x)) ""))

(defun halff:strip-controls (s / i ch out)
  (if (not (and s (= (type s) 'STR))) (setq s (halff:safe-str s)))
  (setq i 1 out "")
  (while (<= i (strlen s))
    (setq ch (ascii (substr s i 1)))
    (if (or (= ch 9) (= ch 32) (>= ch 33))
      (setq out (strcat out (substr s i 1))))
    (setq i (1+ i)))
  out)

(defun halff:trim (s)
  (setq s (halff:nbspace->space s))
  (setq s (halff:strip-controls s))
  (if s (vl-string-trim " \t\r\n" s) ""))

(defun halff:upper (s) (strcase (halff:trim s)))

(defun halff:split (s sep / out p)
  (setq s (halff:trim s) out '())
  (while (and s (> (strlen s) 0) (setq p (vl-string-search sep s)))
    (setq out (cons (halff:trim (substr s 1 p)) out))
    (setq s (substr s (+ p (strlen sep) 1))))
  (if (and s (> (strlen s) 0)) (setq out (cons (halff:trim s) out)))
  (reverse out))

(defun halff:split-many (s / s2 parts)
  (setq s2 (vl-string-translate ";" "," (halff:trim s)))
  (setq parts (halff:split s2 ","))
  (vl-remove-if '(lambda (x) (= (halff:trim x) "")) parts))

(defun halff:setnth (lst idx val / i out)
  (setq i 1 out '())
  (while lst
    (setq out (cons (if (= i idx) val (car lst)) out))
    (setq lst (cdr lst))
    (setq i (1+ i)))
  (reverse out))

(defun halff:pad-to (lst n / out)
  (setq out lst)
  (while (< (length out) n) (setq out (append out (list ""))))
  out)

(defun halff:zeros (n / out)
  (setq out '())
  (repeat n (setq out (cons 0.0 out)))
  (reverse out))

(defun halff:round0 (x)
  (cond
    ((= *HALFF_ROUND_MODE* "UP")
     (if (> x 0.0) (fix (+ x 0.999999)) 0.0))
    ((= *HALFF_ROUND_MODE* "NEAREST")
     (if (> x 0.0) (fix (+ x 0.5)) 0.0))
    (T (fix (+ x 0.5))))
    )

;; ===========================================
;; FORMULA EVALUATOR
;; ===========================================
;; Supported: + - * / ^ ()  variable x  pay-item refs [PayN]
;; [PayN] looks up pay item N in the rawMap built during pass 1.
;; Returns nil on empty/invalid formula, otherwise a float.

(defun halff:formula-tokenize (str / i ch ch2 tok tokens cont)
  (setq i 1 tokens '())
  (while (<= i (strlen str))
    (setq ch (substr str i 1))
    (cond
      ((or (= ch " ") (= ch "\t")) (setq i (1+ i)))
      ((member ch '("+" "-" "*" "/" "^" "(" ")"))
       (setq tokens (append tokens (list ch)) i (1+ i)))
      ((or (= ch "x") (= ch "X"))
       (setq tokens (append tokens (list "x")) i (1+ i)))
      ((= ch "[")
       (setq tok "[" i (1+ i))
       (while (and (<= i (strlen str)) (/= (substr str i 1) "]"))
         (setq tok (strcat tok (substr str i 1)) i (1+ i)))
       (setq tokens (append tokens (list (strcat tok "]"))) i (1+ i)))
      ((or (and (>= (ascii ch) 48) (<= (ascii ch) 57)) (= ch "."))
       (setq tok ch i (1+ i) cont T)
       (while (and cont (<= i (strlen str)))
         (setq ch2 (substr str i 1))
         (if (or (and (>= (ascii ch2) 48) (<= (ascii ch2) 57)) (= ch2 "."))
           (setq tok (strcat tok ch2) i (1+ i))
           (setq cont nil)))
       (setq tokens (append tokens (list tok))))
      (T (setq i (1+ i)))))
  tokens)

(defun halff:fp-primary (toks xval qmap / val paykey pn pair)
  (cond
    ((null toks) (cons 0.0 nil))
    ((= (car toks) "x")
     (cons (float xval) (cdr toks)))
    ((= (car toks) "(")
     (setq val (halff:fp-expr (cdr toks) xval qmap))
     (cons (car val)
           (if (and (cdr val) (= (car (cdr val)) ")"))
             (cdr (cdr val)) (cdr val))))
    ((and (> (strlen (car toks)) 2) (= (substr (car toks) 1 1) "["))
     ;; Extract pay number from [PayN] or [N]
     (setq paykey (substr (car toks) 2 (- (strlen (car toks)) 2)))
     (if (and (>= (strlen paykey) 4)
              (= (strcase (substr paykey 1 3)) "PAY"))
       (setq paykey (substr paykey 4)))
     (setq pn (fix (atof paykey)))
     (setq pair (vl-some '(lambda (p)
                             (if (= (fix (atof (car p))) pn) p nil))
                          qmap))
     (cons (if pair (float (cadr pair)) 0.0) (cdr toks)))
    (T (cons (atof (car toks)) (cdr toks)))))

(defun halff:fp-unary (toks xval qmap / res)
  (if (and toks (= (car toks) "-"))
    (progn (setq res (halff:fp-unary (cdr toks) xval qmap))
           (cons (- (float (car res))) (cdr res)))
    (halff:fp-primary toks xval qmap)))

(defun halff:fp-power (toks xval qmap / res base)
  (setq res (halff:fp-unary toks xval qmap)
        base (car res) toks (cdr res))
  (if (and toks (= (car toks) "^"))
    (progn (setq res (halff:fp-power (cdr toks) xval qmap))
           (cons (expt (float base) (float (car res))) (cdr res)))
    (cons base toks)))

(defun halff:fp-term (toks xval qmap / res lv rv op)
  (setq res (halff:fp-power toks xval qmap)
        lv (car res) toks (cdr res))
  (while (and toks (member (car toks) '("*" "/")))
    (setq op (car toks) toks (cdr toks)
          res (halff:fp-power toks xval qmap)
          rv (car res) toks (cdr res))
    (setq lv (if (= op "*") (* (float lv) (float rv))
               (if (/= rv 0.0) (/ (float lv) (float rv)) 0.0))))
  (cons lv toks))

(defun halff:fp-expr (toks xval qmap / res lv rv op)
  (setq res (halff:fp-term toks xval qmap)
        lv (car res) toks (cdr res))
  (while (and toks (member (car toks) '("+" "-")))
    (setq op (car toks) toks (cdr toks)
          res (halff:fp-term toks xval qmap)
          rv (car res) toks (cdr res))
    (setq lv (if (= op "+") (+ (float lv) (float rv))
               (- (float lv) (float rv)))))
  (cons lv toks))

(defun halff:formula-eval (formula xval qmap / toks res)
  (if (= (halff:trim formula) "") nil
    (progn
      (setq toks (halff:formula-tokenize formula))
      (if toks
        (progn (setq res (halff:fp-expr toks xval qmap))
               (max 0.0 (float (car res))))
        nil))))

;; ===========================================
;; CSV FUNCTIONS
;; ===========================================

(defun halff:csv-escape (s)
  (if (or (vl-string-search "," s) (vl-string-search "\"" s) (vl-string-search "\n" s))
    (strcat "\"" (vl-string-translate "\"" "\"\"" s) "\"")
    s))

(defun halff:csv-line (vals / out)
  (setq out "")
  (foreach v vals
    (setq out (strcat out (if (= out "") "" ",") (halff:csv-escape (halff:safe-str v))))
    )
  out)

(defun halff:csv-write-lines (path lines / f)
  (setq f (open path "w"))
  (if f
    (progn
      (foreach ln lines (write-line ln f))
      (close f)
      T)
    nil))

;; ===========================================
;; VLA OBJECT NAME MAPPING
;; ===========================================

(defun halff:vla-name->dxf (vlaname / vn)
  (setq vn (strcase vlaname))
  (cond
    ((= vn "ACDBLINE") "LINE")
    ((= vn "ACDBLWPOLYLINE") "LWPOLYLINE")
    ((= vn "ACDBPOLYLINE") "POLYLINE")
    ((= vn "ACDB2DPOLYLINE") "POLYLINE")
    ((= vn "ACDB3DPOLYLINE") "POLYLINE")
    ((= vn "ACDBARC") "ARC")
    ((= vn "ACDBHATCH") "HATCH")
    ((= vn "ACDBCIRCLE") "CIRCLE")
    ((= vn "ACDBSPLINE") "SPLINE")
    ((= vn "ACDBINSERT") "INSERT")
    ((= vn "ACDBBLOCKREFERENCE") "INSERT")
    (T vlaname)))

;; ===========================================
;; SEGMENT INTERSECTION & CROSSING DETECTION
;; ===========================================

(defun halff:seg-intersect? (a b c d / ax ay bx by cx cy dx dy den ua ub)
  (setq ax (car a) ay (cadr a))
  (setq bx (car b) by (cadr b))
  (setq cx (car c) cy (cadr c))
  (setq dx (car d) dy (cadr d))
  (setq den (- (* (- bx ax) (- dy cy)) (* (- by ay) (- dx cx))))
  (if (equal den 0.0 1e-12)
    nil
    (progn
      (setq ua (/ (- (* (- dx cx) (- ay cy)) (* (- dy cy) (- ax cx))) den))
      (setq ub (/ (- (* (- bx ax) (- ay cy)) (* (- by ay) (- ax cx))) den))
      (and (<= 0.0 ua) (<= ua 1.0) (<= 0.0 ub) (<= ub 1.0))))
      )

(defun halff:seg-intersect-pt (a b c d / ax ay bx by cx cy dx dy den ua)
  (setq ax (car a) ay (cadr a))
  (setq bx (car b) by (cadr b))
  (setq cx (car c) cy (cadr c))
  (setq dx (car d) dy (cadr d))
  (setq den (- (* (- bx ax) (- dy cy)) (* (- by ay) (- dx cx))))
  (if (equal den 0.0 1e-12)
    nil
    (progn
      (setq ua (/ (- (* (- dx cx) (- ay cy)) (* (- dy cy) (- ax cx))) den))
      (setq ub (/ (- (* (- bx ax) (- ay cy)) (* (- by ay) (- ax cx))) den))
      (if (and (<= 0.0 ua) (<= ua 1.0) (<= 0.0 ub) (<= ub 1.0))
        (list (+ ax (* ua (- bx ax)))
              (+ ay (* ua (- by ay))))
        nil))))

(defun halff:genuine-crossing? (entPts vpPoly / poly i a b p2 c d ixPt distA distB distC distD)
  (setq poly (halff:poly-close vpPoly))
  (setq i 0)
  (while (< i (1- (length entPts)))
    (setq a (nth i entPts))
    (setq b (nth (1+ i) entPts))
    (setq p2 0)
    (while (< p2 (1- (length poly)))
      (setq c (nth p2 poly))
      (setq d (nth (1+ p2) poly))
      (setq ixPt (halff:seg-intersect-pt a b c d))
      (if ixPt
        (progn
          (setq distC (distance ixPt c))
          (setq distD (distance ixPt d))
          (setq distA (distance ixPt a))
          (setq distB (distance ixPt b))
          (if (and (> distC *HALFF_EXCLUSION_DIST*)
                   (> distD *HALFF_EXCLUSION_DIST*)
                   (> distA *HALFF_EXCLUSION_DIST*)
                   (> distB *HALFF_EXCLUSION_DIST*))
            (progn (setq i 999999) (setq p2 999999))
            (setq p2 (1+ p2))))
        (setq p2 (1+ p2))))
    (if (< i 999999) (setq i (1+ i))))
  (= i 999999))

(defun halff:entity-crosses-poly? (pts poly / p2 i a b c d res)
  (setq res nil)
  (setq poly (halff:poly-close poly))
  (setq pts (halff:poly-close pts))
  (setq i 0)
  (while (< i (1- (length pts)))
    (setq a (nth i pts))
    (setq b (nth (1+ i) pts))
    (setq p2 0)
    (while (< p2 (1- (length poly)))
      (setq c (nth p2 poly))
      (setq d (nth (1+ p2) poly))
      (if (halff:seg-intersect? a b c d)
        (progn (setq i 999999 p2 999999) (setq res T)))
      (setq p2 (1+ p2)))
    (setq i (1+ i)))
  res)

(defun halff:vla-get-entity-pts (vlaObj / oname coords pts n i j out sp ep cen r a1 a2 am p1 p2 p3 mn mx xmin ymin xmax ymax dx dy)
  (setq oname (strcase (vla-get-ObjectName vlaObj)))
  (cond
    ((= oname "ACDBLINE")
     (setq sp (halff:variant->list (vla-get-StartPoint vlaObj)))
     (setq ep (halff:variant->list (vla-get-EndPoint vlaObj)))
     (list (list (car sp) (cadr sp)) (list (car ep) (cadr ep))))

    ((= oname "ACDBLWPOLYLINE")
     (setq coords (halff:variant->list (vla-get-Coordinates vlaObj)))
     (setq n (/ (length coords) 2))
     (setq i 0 out '())
     (while (< i n)
       (setq out (cons (list (nth (* 2 i) coords) (nth (1+ (* 2 i)) coords)) out))
       (setq i (1+ i)))
     (reverse out))

    ((= oname "ACDBPOLYLINE")
     (setq out '())
     (setq n (fix (vlax-curve-getEndParam vlaObj)))
     (setq i 0)
     (while (<= i n)
       (setq p2 (vl-catch-all-apply 'vlax-curve-getPointAtParam (list vlaObj (float i))))
       (if (not (vl-catch-all-error-p p2))
         (setq out (cons (list (car p2) (cadr p2)) out)))
       (setq i (1+ i)))
     (if out (reverse out) nil))

    ((= oname "ACDBARC")
     (setq cen (halff:variant->list (vla-get-Center vlaObj)))
     (setq r (vla-get-Radius vlaObj))
     (setq a1 (vla-get-StartAngle vlaObj))
     (setq a2 (vla-get-EndAngle vlaObj))
     (setq am (/ (+ a1 a2) 2.0))
     (setq p1 (list (+ (car cen) (* r (cos a1))) (+ (cadr cen) (* r (sin a1))))
     )
     (setq p2 (list (+ (car cen) (* r (cos am))) (+ (cadr cen) (* r (sin am))))
     )
     (setq p3 (list (+ (car cen) (* r (cos a2))) (+ (cadr cen) (* r (sin a2))))
     )
     (list p1 p2 p3))

    ((= oname "ACDBHATCH")
     (vla-GetBoundingBox vlaObj 'mn 'mx)
     (setq mn (halff:variant->list mn))
     (setq mx (halff:variant->list mx))
     (setq out '())
     (setq xmin (car mn)) (setq ymin (cadr mn))
     (setq xmax (car mx)) (setq ymax (cadr mx))
     (setq dx (/ (- xmax xmin) 20.0))
     (setq dy (/ (- ymax ymin) 20.0))
     (setq i 0)
     (while (<= i 20)
       (setq out (cons (list (+ xmin (* i dx)) ymax) out))
       (setq out (cons (list (+ xmin (* i dx)) ymin) out))
       (setq out (cons (list xmin (+ ymin (* i dy))) out))
       (setq out (cons (list xmax (+ ymin (* i dy))) out))
       (setq i (1+ i)))
     (setq i 4)
     (while (<= i 16)
       (setq j 4)
       (while (<= j 16)
         (setq out (cons (list (+ xmin (* i dx)) (+ ymin (* j dy))) out))
         (setq j (+ j 4)))
       (setq i (+ i 4)))
     out)

    (T nil)))

(defun halff:pts-in-vp? (pts vpPoly / found p)
  (if (not pts)
    nil
    (progn
      (setq found nil)
      (while (and pts (not found))
        (setq p (car pts))
        (if (halff:pt-in-poly p vpPoly) (setq found T))
        (setq pts (cdr pts)))
      found)))

(defun halff:vla-ent-has-points-in-vp? (vlaObj vpPoly / oname entPts p found)
  (setq oname (strcase (vla-get-ObjectName vlaObj)))
  (setq entPts (halff:vla-get-entity-pts vlaObj))
  (if (not entPts)
    nil
    (progn
      (setq found nil)
      (while (and entPts (not found))
        (setq p (car entPts))
        (if (halff:pt-in-poly p vpPoly) (setq found T))
        (setq entPts (cdr entPts)))
      found)))

(defun halff:filter-clear-points (entPts allVPs / clearPts p isNearAnyBoundary i vpPolyClosed minDist j d)
  (setq clearPts '())
  (foreach p entPts
    (setq isNearAnyBoundary nil)
    (setq i 0)
    (while (and (< i (length allVPs)) (not isNearAnyBoundary))
      (setq vpPolyClosed (halff:poly-close (cadr (nth i allVPs))))
      (setq minDist 999999.0)
      (setq j 0)
      (while (< j (1- (length vpPolyClosed)))
        (setq d (halff:pt-to-seg-dist p (nth j vpPolyClosed) (nth (1+ j) vpPolyClosed)))
        (if (< d minDist) (setq minDist d))
        (setq j (1+ j)))
      (if (< minDist *HALFF_EXCLUSION_DIST*) (setq isNearAnyBoundary T))
      (setq i (1+ i)))
    (if (not isNearAnyBoundary) (setq clearPts (cons p clearPts))))
  clearPts)

(defun halff:vla-ent-crosses-vp? (vlaObj vpPoly / oname entPts anyIn anyOut countIn countOut totalPts pctIn pctOut p minDist j d vpPolyClosed clearPts)
  (setq oname (strcase (vla-get-ObjectName vlaObj)))
  (setq entPts (halff:vla-get-entity-pts vlaObj))
  (if (not entPts)
    nil
    (progn
      (setq totalPts (length entPts))
      (if (member oname '("ACDBHATCH" "ACDBLWPOLYLINE"))
        (progn
          (setq vpPolyClosed (halff:poly-close vpPoly))
          (setq clearPts '())
          (foreach p entPts
            (setq minDist 999999.0)
            (setq j 0)
            (while (< j (1- (length vpPolyClosed)))
              (setq d (halff:pt-to-seg-dist p (nth j vpPolyClosed) (nth (1+ j) vpPolyClosed)))
              (if (< d minDist) (setq minDist d))
              (setq j (1+ j)))
            (if (> minDist *HALFF_EXCLUSION_DIST*)
              (setq clearPts (cons p clearPts))))
          (setq entPts clearPts)
          (setq totalPts (length entPts))))
      (if (< totalPts 2)
        nil
        (progn
          (setq anyIn nil anyOut nil)
          (setq countIn 0 countOut 0)
          (foreach p entPts
            (if (halff:pt-in-poly p vpPoly)
              (progn (setq anyIn T) (setq countIn (1+ countIn)))
              (progn (setq anyOut T) (setq countOut (1+ countOut))))
              )
          (if (and anyIn anyOut)
            (progn
              (setq pctIn (/ (* countIn 100.0) totalPts))
              (setq pctOut (/ (* countOut 100.0) totalPts))
              (if (= oname "ACDBHATCH")
                (and anyIn anyOut)
                (if (>= totalPts 2)
                  (halff:entity-crosses-poly? entPts vpPoly)
                  nil)))
            nil))))
            ))

(defun halff:pt-to-seg-dist (pt p1 p2 / px py x1 y1 x2 y2 dx dy dsq param cx cy)
  (setq px (car pt) py (cadr pt))
  (setq x1 (car p1) y1 (cadr p1))
  (setq x2 (car p2) y2 (cadr p2))
  (setq dx (- x2 x1) dy (- y2 y1))
  (setq dsq (+ (* dx dx) (* dy dy)))
  (if (< dsq 0.0001)
    (distance pt p1)
    (progn
      (setq param (/ (+ (* (- px x1) dx) (* (- py y1) dy)) dsq))
      (if (< param 0.0) (setq param 0.0))
      (if (> param 1.0) (setq param 1.0))
      (setq cx (+ x1 (* param dx)))
      (setq cy (+ y1 (* param dy)))
      (distance pt (list cx cy))))
      )

;; ===========================================
;; VIEWPORT FUNCTIONS
;; ===========================================

(defun halff:pt->str (pt)
  (strcat (rtos (car pt) 2 8) "," (rtos (cadr pt) 2 8)))

(defun halff:str->pt (s / p)
  (setq p (vl-string-search "," s))
  (list (atof (substr s 1 p))
        (atof (substr s (+ p 2))))
        )

(defun halff:get-lwpoly-pts (vlaObj / coords pts i n out)
  (setq coords (halff:variant->list (vla-get-Coordinates vlaObj)))
  (setq pts coords)
  (setq n (/ (length pts) 2))
  (setq i 0 out '())
  (while (< i n)
    (setq out (cons (list (nth (* 2 i) pts) (nth (1+ (* 2 i)) pts)) out))
    (setq i (1+ i)))
  (reverse out))

(defun halff:poly-close (pts)
  (if (and pts (/= (car pts) (last pts)))
    (append pts (list (car pts)))
    pts))

(defun halff:pt-in-poly (pt poly / x y inside i j xi yi xj yj)
  (setq x (car pt) y (cadr pt))
  (setq inside nil)
  (setq i 0 j (1- (length poly)))
  (while (< i (length poly))
    (setq xi (car (nth i poly)) yi (cadr (nth i poly)))
    (setq xj (car (nth j poly)) yj (cadr (nth j poly)))
    (if (and (/= yi yj)
             (<= (min yi yj) y)
             (<  y (max yi yj))
             (<  x (+ xi (* (/ (- y yi) (- yj yi)) (- xj xi))))
             )
      (setq inside (not inside)))
    (setq j i)
    (setq i (1+ i)))
  inside)

(defun halff:vp-prep (vplist / out nm pts bb xs ys)
  (setq out '())
  (foreach vp vplist
    (setq nm (car vp) pts (cdr vp))
    (setq xs (mapcar 'car pts) ys (mapcar 'cadr pts))
    (setq bb (list (apply 'min xs) (apply 'min ys) (apply 'max xs) (apply 'max ys)))
    (setq out (append out (list (list nm pts bb))))
    )
  out)

(defun halff:vla-bbox (e / o mn mx a b res)
  (setq res (vl-catch-all-apply
    '(lambda ()
      (setq o (vlax-ename->vla-object e))
      (vla-getboundingbox o 'mn 'mx)
      (setq a (halff:variant->list mn))
      (setq b (halff:variant->list mx))
      (list (car a) (cadr a) (car b) (cadr b))))
      )
  (if (vl-catch-all-error-p res) nil res))

(defun halff:bbox-overlap (b1 b2)
  (and (<= (nth 0 b1) (nth 2 b2))
       (>= (nth 2 b1) (nth 0 b2))
       (<= (nth 1 b1) (nth 3 b2))
       (>= (nth 3 b1) (nth 1 b2))))

(defun halff:pt-in-bbox (pt bb)
  (and (>= (car pt) (nth 0 bb)) (<= (car pt) (nth 2 bb))
       (>= (cadr pt) (nth 1 bb)) (<= (cadr pt) (nth 3 bb))))

(defun halff:vla-ent-in-vp (vlaObj vp / pts vbb mn mx ebb cx cy center res mnList mxList)
  (setq pts (cadr vp))
  (setq vbb (caddr vp))
  (setq res (vl-catch-all-apply
    '(lambda ()
      (vla-getboundingbox vlaObj 'mn 'mx)
      (setq mnList (halff:variant->list mn))
      (setq mxList (halff:variant->list mx))
      (list (car mnList) (cadr mnList) (car mxList) (cadr mxList))))
      )
  (if (vl-catch-all-error-p res)
    nil
    (progn
      (setq ebb res)
      (if (not (halff:bbox-overlap ebb vbb))
        nil
        (progn
          (setq cx (/ (+ (nth 0 ebb) (nth 2 ebb)) 2.0))
          (setq cy (/ (+ (nth 1 ebb) (nth 3 ebb)) 2.0))
          (setq center (list cx cy))
          (halff:pt-in-poly center pts))))
          ))

;; ===========================================
;; GEOMETRY FUNCTIONS
;; ===========================================

(defun halff:length (e / o res)
  (setq res (vl-catch-all-apply
    '(lambda ()
      (setq o (vlax-ename->vla-object e))
      (if (= (vla-get-ObjectName o) "AcDbCircle")
        (* 2.0 pi (vla-get-Radius o))
        (vlax-curve-getDistAtParam e (vlax-curve-getEndParam e))))
        ))
  (if (vl-catch-all-error-p res) 0.0 res))

(defun halff:area (e / o res)
  (setq res (vl-catch-all-apply
    '(lambda ()
      (setq o (vlax-ename->vla-object e))
      (if (vlax-property-available-p o 'Area) (vla-get-Area o) 0.0))))
  (if (vl-catch-all-error-p res) 0.0 res))

;; ===========================================
;; UNIT CONVERSION
;; ===========================================

(defun halff:mode-from-unit (u / uu)
  (setq uu (halff:upper u))
  (cond
    ((member uu '("EA" "EACH" "COUNT")) "COUNT")
    ((member uu '("LF" "LNFT" "FT" "FEET")) "LENGTH_FT")
    ((member uu '("LY" "YD" "YARD" "YARDS")) "LENGTH_YD")
    ((member uu '("SF" "SQFT" "FT2" "SQFT.")) "AREA_FT2")
    ((member uu '("SY" "SQYD" "SQYDS" "YD2" "SYD" "SQYD.")) "AREA_YD2")
    (T "COUNT")))

(defun halff:feet-scale ( / iu)
  (setq iu (getvar "INSUNITS"))
  (cond
    ((member iu '(0 2 21)) 1.0)
    ((= iu 1) (/ 1.0 12.0))
    ((= iu 6) 3.280839895)
    ((= iu 5) 0.03280839895)
    ((= iu 4) 0.003280839895)
    (T 1.0)))

(defun halff:convert (q mode fs)
  (cond
    ((= mode "COUNT") q)
    ((= mode "LENGTH_FT") (* q fs))
    ((= mode "LENGTH_YD") (/ (* q fs) 3.0))
    ((= mode "AREA_FT2") (* q fs fs))
    ((= mode "AREA_YD2") (/ (* q fs fs) 9.0))
    (T q)))

(defun halff:apply-mult (q m / mm)
  (setq m (halff:trim m))
  (if (= m "") q
    (progn (setq mm (atof m)) (if (= mm 0.0) q (* q mm))))
    )

;; ===========================================
;; TEXT CONTENT SEARCH HELPERS
;; ===========================================

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

(defun halff:text-contains? (needle haystack / n h)
  (setq n (strcase (halff:normalize-ws needle))
        h (strcase (halff:normalize-ws haystack)))
  (not (null (vl-string-search n h))))

(defun halff:text-obj-types (objStr / types out u searchAll)
  ;; Map the OBJECT column value to a list of uppercase VLA ObjectName
  ;; strings used to filter entities during text-content search.
  ;; Returns nil when any unrecognised type is present, which tells the
  ;; caller to try ALL entities (needed for Civil 3D label types whose
  ;; DXF entity name is not a standard AutoCAD name).
  (setq types     (halff:split-many objStr)
        out       '()
        searchAll nil)
  (foreach tp types
    (if (not searchAll)
      (progn
        (setq u (halff:upper tp))
        (cond
          ((= u "TEXT")
           (setq out (cons "ACDBTEXT" out)))
          ((= u "MTEXT")
           (setq out (cons "ACDBMTEXT" out)))
          ((member u '("MLEADER" "MULTILEADER"))
           (setq out (cons "ACDBMLEADER" out)))
          (T
           ;; Unrecognised (e.g. "Alignment Station Offset Label") →
           ;; search every entity in the file for text content.
           (setq searchAll T))))
           ))
  (if searchAll nil (reverse out)))

;; ===========================================
;; SELECTION MAPPING
;; ===========================================

(defun halff:obj->dxf (o / u)
  (setq u (halff:upper o))
  (cond
    ((= u "HATCH") "HATCH")
    ((= u "LINE") "LINE")
    ((member u '("POLYLINE" "PLINE" "P-LINE")) "LWPOLYLINE,POLYLINE")
    ((= u "LWPOLYLINE") "LWPOLYLINE")
    ((= u "ARC") "ARC")
    ((= u "SPLINE") "SPLINE")
    ((= u "CIRCLE") "CIRCLE")
    ((member u '("INSERT" "BLOCK" "BLOCKS" "BLOCK REFERENCE")) "INSERT")
    (T u)))

(defun halff:types->dxf0 (types / out cur part tmp)
  (setq out '())
  (while types
    (setq cur (halff:obj->dxf (car types)))
    (setq tmp (halff:split cur ","))
    (while tmp
      (setq part (halff:upper (car tmp)))
      (if (and (/= part "") (not (member part out))) (setq out (cons part out)))
      (setq tmp (cdr tmp)))
    (setq types (cdr types)))
  (setq out (reverse out))
  (setq cur "")
  (while out
    (setq cur (strcat cur (car out) (if (cdr out) "," "")))
    (setq out (cdr out)))
  cur)

;; ===========================================
;; FILTERING
;; ===========================================

(defun halff:ent-pass-filters (e ltype hatchpat / ed etype lt pat)
  (setq ed (entget e))
  (setq etype (strcase (cdr (assoc 0 ed))))
  (setq lt (cdr (assoc 6 ed)))
  (if lt (setq lt (strcase lt)) (setq lt ""))
  (setq pat (if (= etype "HATCH") (strcase (cdr (assoc 2 ed))) ""))
  (and
    (or (= (halff:trim ltype) "") (= lt (strcase (halff:trim ltype))))
    (or (= (halff:trim hatchpat) "") (and (= etype "HATCH") (= pat (strcase (halff:trim hatchpat))))
    )))

;; ===========================================
;; MEASUREMENT & QUANTIFICATION
;; ===========================================

(defun halff:ent-measure (e mode fs / one)
  (setq one
    (vl-catch-all-apply
      '(lambda ()
        (cond
          ((= mode "COUNT") 1.0)
          ((member mode '("LENGTH_FT" "LENGTH_YD")) (halff:convert (halff:length e) mode fs))
          ((member mode '("AREA_FT2" "AREA_YD2")) (halff:convert (halff:area e) mode fs))
          (T 0.0))))
          )
  (if (vl-catch-all-error-p one) 0.0
    (if (and *HALFF_ROUND_EACH* (not (= mode "COUNT"))) (halff:round0 one) one)))

(defun halff:log-failure (payitem filepath layer objtype vpnames reason)
  (setq *HALFF_FAILURE_LOG*
    (append *HALFF_FAILURE_LOG*
      (list (list payitem filepath layer objtype
                 (if (listp vpnames)
                   (if vpnames
                     (apply 'strcat
                       (cons (car vpnames)
                         (mapcar '(lambda (vp) (strcat "," vp)) (cdr vpnames))))
                     "N/A")
                   vpnames)
                 reason))))
                 )

(defun halff:sum-per-vps (ss mode fs ltype hatchpat vps payitem filepath / i e meas out outside hit j vpname)
  (setq out (halff:zeros (length vps)))
  (setq outside 0.0)
  (setq i 0)
  (while (and ss (< i (sslength ss)))
    (setq e (ssname ss i))
    (if (halff:ent-pass-filters e ltype hatchpat)
      (progn
        (setq meas (halff:ent-measure e mode fs))
        (if (= meas 0.0)
          (halff:log-failure payitem filepath (cdr (assoc 8 (entget e)))
                            (cdr (assoc 0 (entget e))) "N/A" "Geometry calculation failed"))
        (setq hit nil)
        (if (> meas 0.0)
          (progn
            (setq j 0)
            (while (< j (length vps))
              (if (halff:ent-in-vp e (nth j vps))
                (progn
                  (setq out (halff:setnth out (1+ j) (+ (nth j out) meas)))
                  (setq hit T)))
              (setq j (1+ j)))
            (if (not hit)
              (progn
                (setq outside (+ outside meas))
                (halff:log-failure payitem filepath (cdr (assoc 8 (entget e)))
                                  (cdr (assoc 0 (entget e))) "OUTSIDE" "Entity outside all viewports"))))
                                  ))
      (halff:log-failure payitem filepath (cdr (assoc 8 (entget e)))
                        (cdr (assoc 0 (entget e))) "N/A" "Failed filter (linetype or pattern)"))
    (setq i (1+ i)))
  (list (apply '+ out) out outside))

;; ===========================================
;; EXCEL FUNCTIONS
;; ===========================================

(defun halff:xl-open-ro (xlsx / xl wb res wbs)
  (setq res (vl-catch-all-apply
    '(lambda ()
      (setq xl (vlax-get-or-create-object "Excel.Application"))
      (vlax-put-property xl 'Visible :vlax-false)
      (vlax-put-property xl 'DisplayAlerts :vlax-false)
      (vlax-put-property xl 'ScreenUpdating :vlax-false)
      (setq wbs (vlax-get-property xl 'Workbooks))
      (setq wb (vlax-invoke-method wbs 'Open xlsx))
      (list xl wb))))
  (if (vl-catch-all-error-p res) res (if res res (vl-catch-all-error "Excel open failed"))))

(defun halff:xl-close (xl wb)
  (if wb (vl-catch-all-apply 'vlax-invoke-method (list wb 'Close 0)))
  (if xl (progn (vl-catch-all-apply 'vlax-invoke-method (list xl 'Quit))
                (vl-catch-all-apply 'vlax-release-object (list xl))))
                )

(defun halff:ws-active (wb)
  (vlax-get-property wb 'ActiveSheet))

(defun halff:used-rows (ws / ur)
  (setq ur (vlax-get-property ws 'UsedRange))
  (vlax-get-property (vlax-get-property ur 'Rows) 'Count))

(defun halff:used-cols (ws / ur)
  (setq ur (vlax-get-property ws 'UsedRange))
  (vlax-get-property (vlax-get-property ur 'Columns) 'Count))

(defun halff:get-cell (ws row col / cells range val)
  (setq cells (vlax-get-property ws 'Cells))
  (setq range (vlax-get-property cells 'Item row col))
  (if (= (type range) 'VARIANT)
    (setq range (vlax-variant-value range)))
  (setq val (vlax-get-property range 'Value))
  (if (= (type val) 'VARIANT)
    (vlax-variant-value val)
    val))

(defun halff:get-row-string (ws row col / v res)
  (setq res (vl-catch-all-apply 'halff:get-cell (list ws row col)))
  (if (vl-catch-all-error-p res)
    ""
    (progn
      (setq v res)
      (if v (halff:trim (vl-princ-to-string (halff:variant->list v))) ""))))

(defun halff:headers (ws / maxc hdr i)
  (setq maxc (halff:used-cols ws))
  (setq hdr '() i 1)
  (while (<= i maxc)
    (setq hdr (append hdr (list (halff:upper (halff:get-row-string ws 1 i))))
    )
    (setq i (1+ i)))
  hdr)

(defun halff:col (hdr name / i found)
  (setq i 1 found nil)
  (while (and (<= i (length hdr)) (not found))
    (if (= (nth (1- i) hdr) (halff:upper name))
      (setq found i))
    (setq i (1+ i)))
  found)

;; ===========================================
;; SIDECAR FILE FUNCTIONS
;; ===========================================

(defun halff:sidecar-path ()
  (strcat (getvar "DWGPREFIX") (getvar "DWGNAME") ".halffqto"))

(defun halff:file-read-first-line (fp / f line)
  (if (findfile fp)
    (progn (setq f (open fp "r")) (setq line (if f (read-line f) nil)) (if f (close f)) line)
    nil))

(defun halff:file-write-line (fp line / f)
  (setq f (open fp "w"))
  (if f (progn (write-line line f) (close f) T) nil))

(defun halff:get-mapping-path () (halff:file-read-first-line (halff:sidecar-path)))
(defun halff:set-mapping-path (p) (halff:file-write-line (halff:sidecar-path) p))

(defun halff:get-vp-path ()
  (halff:file-read-first-line (strcat (getvar "DWGPREFIX") (getvar "DWGNAME") ".halffvp")))
(defun halff:set-vp-path (p)
  (halff:file-write-line (strcat (getvar "DWGPREFIX") (getvar "DWGNAME") ".halffvp") p))

;; ===========================================
;; DBX FUNCTIONS FOR EXTERNAL FILE ACCESS
;; ===========================================

(defun halff:acadver-major (/ v p)
  (setq v (getvar "ACADVER"))
  (setq p (vl-string-search "." v))
  (if p (atoi (substr v 1 p)) 0))

(defun halff:dbx-progids (/ maj)
  (setq maj (halff:acadver-major))
  (list "ObjectDBX.AxDbDocument"
        (strcat "ObjectDBX.AxDbDocument." (itoa maj))
        (strcat "ObjectDBX.AxDbDocument." (itoa (1- maj)))
        (strcat "ObjectDBX.AxDbDocument." (itoa (+ maj 1))))
        )

(defun halff:dbx-create (/ acad pid obj err err2)
  (setq acad (vlax-get-acad-object))
  (setq obj nil)
  (foreach pid (halff:dbx-progids)
    (if (not obj)
      (progn
        (setq err (vl-catch-all-apply 'vlax-create-object (list pid)))
        (if (not (vl-catch-all-error-p err))
          (setq obj err)
          (progn
            (setq err2 (vl-catch-all-apply 'vla-getInterfaceObject (list acad pid)))
            (if (not (vl-catch-all-error-p err2))
              (setq obj err2))))
              )))
  obj)

;; ===========================================
;; VIEWPORT COMMANDS
;; ===========================================

(defun c:QVPDEF (/ ss i en obj name pts out path vstr)
  (prompt "\nSelect ALL VP boundary polylines, then press ENTER.")
  (setq ss (ssget '((0 . "LWPOLYLINE"))))
  (if (not ss)
    (prompt "\nNothing selected.")
    (progn
      (setq *HALFF_VP_DEFS* nil)
      (setq i 0)
      (while (< i (sslength ss))
        (setq en (ssname ss i))
        (setq obj (vlax-ename->vla-object en))
        (setq pts (halff:get-lwpoly-pts obj))
        (setq name (strcat "VP" (itoa (1+ i))))
        (setq *HALFF_VP_DEFS* (cons (cons name pts) *HALFF_VP_DEFS*))
        (setq i (1+ i)))
      (setq *HALFF_VP_DEFS* (reverse *HALFF_VP_DEFS*))
      (setq path (getfiled "Save VP Definitions CSV"
                          (strcat (getvar "DWGPREFIX") "vp_defs.csv") "csv" 1))
      (if path
        (progn
          (halff:set-vp-path path)
          (setq out (list "VP_NAME,VERTS"))
          (foreach kv *HALFF_VP_DEFS*
            (setq name (car kv))
            (setq pts (cdr kv))
            (setq vstr "")
            (foreach p pts
              (setq vstr (strcat vstr (if (= vstr "") "" ";") (halff:pt->str p))))
            (setq out (append out (list (strcat name ",\"" vstr "\""))))
            )
          (halff:csv-write-lines path out)
          (prompt (strcat "\nSaved " (itoa (length *HALFF_VP_DEFS*)) " VP defs: " path)))
        (prompt "\nCancelled."))))
  (princ))

(defun c:QVPLOAD (/ path f ln parts name vstr pts tok)
  (setq path (halff:get-vp-path))
  (if (not path)
    (setq path (getfiled "Load VP Definitions CSV"
                        (strcat (getvar "DWGPREFIX") "vp_defs.csv") "csv" 0)))
  (if (not path)
    (prompt "\nCancelled.")
    (progn
      (setq f (open path "r"))
      (if (not f)
        (prompt (strcat "\nERROR: Cannot open " path))
        (progn
          (setq *HALFF_VP_DEFS* nil)
          (setq ln (read-line f))
          (while (setq ln (read-line f))
            (setq parts (halff:split ln ",\""))
            (setq name (car parts))
            (setq vstr (cadr parts))
            (if (and vstr (> (strlen vstr) 0))
              (setq vstr (substr vstr 1 (1- (strlen vstr))))
              )
            (setq pts '())
            (foreach tok (halff:split vstr ";")
              (setq pts (cons (halff:str->pt tok) pts)))
            (setq pts (reverse pts))
            (setq *HALFF_VP_DEFS* (cons (cons name pts) *HALFF_VP_DEFS*)))
          (close f)
          (setq *HALFF_VP_DEFS* (reverse *HALFF_VP_DEFS*))
          (halff:set-vp-path path)
          (prompt (strcat "\nLoaded " (itoa (length *HALFF_VP_DEFS*)) " VP defs from: " path))))
          ))
  (princ))

(defun c:QVPSET (/ cur p)
  (setq cur (halff:get-vp-path))
  (if (and cur (/= cur "")) (princ (strcat "\nCurrent VP definitions: " cur)))
  (setq p (getfiled "Select VP Definitions CSV" (if cur cur (strcat (getvar "DWGPREFIX") "vp_defs.csv")) "csv" 0))
  (if (and p (/= p ""))
    (progn
      (halff:set-vp-path p)
      (princ (strcat "\nVP definitions path set: " p))
      (princ "\nRun QVPLOAD to load the viewports."))
    (princ "\nNo file selected."))
  (princ))

;; ===========================================
;; MAIN QTO FUNCTION - PROCESS ONE DWG FILE
;; (geometry / layer mode — unchanged from v1.22)
;; ===========================================

(defun halff:process-dwg-file (dwgPath payitem unit layer objstr ltype hatchpat mult vps / dbx openRes ms mode fs types tstr found result vpVals dxfTypes ent dxfName entList totalQty vpQtys i qty j hit hitCount firstHitIdx crossesAny vpNamesIn vpNamesCrossing outside curPath isCurrentDwg entPts isFail objType)
  (setq curPath (strcat (getvar "DWGPREFIX") (getvar "DWGNAME")))
  (setq isCurrentDwg (= (strcase dwgPath) (strcase curPath)))
  (if isCurrentDwg
    (progn
      (princ "\n    Processing CURRENT drawing")
      (setq ms (vla-get-ModelSpace (vla-get-ActiveDocument (vlax-get-acad-object))))
      )
    (progn
      (setq dbx (halff:dbx-create))
      (if (not dbx)
        (progn
          (princ (strcat "\n    ERROR: Cannot create DBX for " dwgPath))
          (halff:log-failure payitem dwgPath layer objstr "N/A" "Cannot create DBX object")
          (list 0.0 (halff:zeros (length vps))))
        (progn
          (setq openRes (vl-catch-all-apply 'vla-open (list dbx dwgPath)))
          (if (vl-catch-all-error-p openRes)
            (progn
              (princ (strcat "\n    ERROR: Cannot open " dwgPath))
              (princ (strcat "\n    " (vl-catch-all-error-message openRes)))
              (halff:log-failure payitem dwgPath layer objstr "N/A"
                                (strcat "Cannot open file: " (vl-catch-all-error-message openRes)))
              (vl-catch-all-apply 'vlax-release-object (list dbx))
              (list 0.0 (halff:zeros (length vps))))
              )
          (if (not (vl-catch-all-error-p openRes))
            (setq ms (vla-get-ModelSpace dbx))))
            )))
  (if (not ms)
    (list 0.0 (halff:zeros (length vps)))
    (progn
      (setq mode (halff:mode-from-unit unit))
      (setq fs (halff:feet-scale))
      (setq types (halff:split-many objstr))
      (setq tstr (halff:types->dxf0 types))
      (setq dxfTypes (halff:split tstr ","))
      (princ (strcat "\n    Processing: " dwgPath))
      (princ (strcat "\n      Layer=" layer " Object=" objstr " Mode=" mode))
      (princ (strcat "\n      DXF Types=" tstr))
      (setq entList '())
      (setq found 0)
      (vlax-for ent ms
        (setq dxfName (halff:vla-name->dxf (vla-get-ObjectName ent)))
        (if (and (member dxfName dxfTypes)
                 (= (strcase (vla-get-Layer ent)) (strcase layer))
                 (or (= (halff:trim ltype) "")
                     (= (strcase (vla-get-Linetype ent)) (strcase (halff:trim ltype))))
                 (or (= (halff:trim hatchpat) "")
                     (/= dxfName "HATCH")
                     (= (strcase (vla-get-PatternName ent)) (strcase (halff:trim hatchpat))))
                     )
          (progn
            (setq entList (cons ent entList))
            (setq found (1+ found))))
            )
      (princ (strcat "\n      Found " (itoa found) " entities"))
      (if (> found 0)
        (progn
          (setq vpQtys (halff:zeros (length vps)))
          (setq totalQty 0.0)
          (setq outside 0.0)
          (foreach ent entList
            (setq isFail nil)
            (setq qty (vl-catch-all-apply
              '(lambda ()
                (cond
                  ((= mode "COUNT") 1.0)
                  ((member mode '("LENGTH_FT" "LENGTH_YD"))
                   (halff:convert
                     (if (= (vla-get-ObjectName ent) "AcDbCircle")
                       (* 2.0 pi (vla-get-Radius ent))
                       (if (= (vla-get-ObjectName ent) "AcDbArc")
                         (* (vla-get-Radius ent)
                            (- (vla-get-EndAngle ent) (vla-get-StartAngle ent)))
                         (vla-get-Length ent)))
                     mode fs))
                  ((member mode '("AREA_FT2" "AREA_YD2"))
                   (halff:convert (vla-get-Area ent) mode fs))
                  (T 0.0))))
                  )
            (if (vl-catch-all-error-p qty)
              (progn
                (setq qty 0.0)
                (setq isFail T)
                (halff:log-failure payitem dwgPath layer
                  (halff:vla-name->dxf (vla-get-ObjectName ent)) "N/A" "Geometry calculation failed"))
              (if (and *HALFF_ROUND_EACH* (not (= mode "COUNT")))
                (setq qty (halff:round0 qty))))
            (if isCurrentDwg
              (halff:remember-entity payitem ent))
            (if (> qty 0.0)
              (progn
                (setq hit nil)
                (setq hitCount 0)
                (setq firstHitIdx -1)
                (setq crossesAny nil)
                (setq vpNamesIn '())
                (setq vpNamesCrossing '())
                (setq entPts (halff:vla-get-entity-pts ent))
                (setq objType (strcase (vla-get-ObjectName ent)))
                (setq i 0)
                (while (< i (length vps))
                  (if (halff:vla-ent-in-vp ent (nth i vps))
                    (progn
                      (if (= firstHitIdx -1) (setq firstHitIdx i))
                      (setq vpNamesIn (append vpNamesIn (list (car (nth i vps))))
                      )
                      (setq hit T)
                      (setq hitCount (1+ hitCount))))
                  (setq i (1+ i)))
                (setq i 0)
                (while (< i (length vps))
                  (if (= objType "ACDBHATCH")
                    (if (halff:vla-ent-crosses-vp? ent (cadr (nth i vps)))
                      (setq vpNamesCrossing (append vpNamesCrossing (list (car (nth i vps))))
                      ))
                    (if (and (/= i firstHitIdx) entPts (>= (length entPts) 2))
                      (if (halff:genuine-crossing? entPts (cadr (nth i vps)))
                        (setq vpNamesCrossing (append vpNamesCrossing (list (car (nth i vps))))
                        ))))
                  (setq i (1+ i)))
                (if (> (length vpNamesCrossing) 0)
                  (progn
                    (setq crossesAny T)
                    (halff:log-failure payitem dwgPath layer
                                       (halff:vla-name->dxf (vla-get-ObjectName ent))
                                       vpNamesCrossing "Entity crosses viewport boundary")))
                (if crossesAny (setq isFail T))
                (cond
                  ((and crossesAny *HALFF_CROSSING_QTY_ZERO*)
                   (setq totalQty (+ totalQty qty))
                   (setq outside (+ outside qty)))
                  ((> hitCount 1)
                   (setq isFail T)
                   (princ (strcat "\n      WARNING: Entity center in " (itoa hitCount) " viewports - excluding"))
                   (setq outside (+ outside qty))
                   (halff:log-failure payitem dwgPath layer
                     (halff:vla-name->dxf (vla-get-ObjectName ent))
                     vpNamesIn "Entity center in multiple viewports"))
                  ((= hitCount 1)
                   (setq vpQtys (halff:setnth vpQtys (1+ firstHitIdx) (+ (nth firstHitIdx vpQtys) qty)))
                   (setq totalQty (+ totalQty qty)))
                  (T
                   (setq isFail T)
                   (setq outside (+ outside qty))
                   (halff:log-failure payitem dwgPath layer
                     (halff:vla-name->dxf (vla-get-ObjectName ent))
                     "OUTSIDE" "Entity outside all viewports")))
                (if (and isCurrentDwg isFail)
                  (halff:remember-fail-entity ent))))
                  )
          (princ (strcat "\n      Total=" (rtos totalQty 2 4)))
          (if (and (not isCurrentDwg) dbx) (vl-catch-all-apply 'vlax-release-object (list dbx)))
          (list (halff:apply-mult totalQty mult)
                (mapcar '(lambda (v) (halff:apply-mult v mult)) vpQtys)))
        (progn
          (princ "\n      No entities found")
          (halff:log-failure payitem dwgPath layer objstr "N/A" "No entities match filters")
          (if (and (not isCurrentDwg) dbx) (vl-catch-all-apply 'vlax-release-object (list dbx)))
          (list 0.0 (halff:zeros (length vps))))
          ))))

;; ===========================================
;; TEXT CONTENT COUNT - PROCESS ONE DWG FILE
;; ===========================================

(defun halff:count-text-in-file
       (dwgPath payitem textContent objStr vps
        / dbx openRes ms onames searchAll obj oname txt
          found totalQty vpQtys hitCount firstHitIdx i inVpRes
          isCurrentDwg curPath en)

  (setq curPath    (strcat (getvar "DWGPREFIX") (getvar "DWGNAME"))
        isCurrentDwg (= (strcase dwgPath) (strcase curPath)))

  ;; Open the DWG — same DBX approach as halff:process-dwg-file
  (if isCurrentDwg
    (progn
      (princ "\n    Processing CURRENT drawing (text-count mode)")
      (setq ms (vla-get-ModelSpace
                 (vla-get-ActiveDocument (vlax-get-acad-object))))
                 )
    (progn
      (setq dbx (halff:dbx-create))
      (if (not dbx)
        (progn
          (princ (strcat "\n    ERROR: Cannot create DBX for " dwgPath))
          (halff:log-failure payitem dwgPath "" objStr "N/A"
                             "Cannot create DBX object")
          (setq ms nil))
        (progn
          (setq openRes (vl-catch-all-apply 'vla-open (list dbx dwgPath)))
          (if (vl-catch-all-error-p openRes)
            (progn
              (princ (strcat "\n    ERROR: Cannot open " dwgPath))
              (halff:log-failure payitem dwgPath "" objStr "N/A"
                                 (strcat "Cannot open: "
                                         (vl-catch-all-error-message openRes)))
              (vl-catch-all-apply 'vlax-release-object (list dbx))
              (setq ms nil))
            (setq ms (vla-get-ModelSpace dbx))))
            )))

  (if (not ms)
    (list 0.0 (halff:zeros (length vps)))
    (progn
      ;; Resolve entity type filter
      (setq onames    (halff:text-obj-types objStr)
            searchAll (null onames))
      (princ (strcat "\n    Processing: " dwgPath))
      (princ (strcat "\n      Text-count: \"" textContent "\""))
      (princ (strcat "\n      Object filter: "
                     (if searchAll
                       "ALL entities (Civil 3D / unrecognised type)"
                       objStr)))

      (setq found    0
            totalQty 0.0
            vpQtys   (halff:zeros (length vps)))

      (vlax-for obj ms
        (setq oname (strcase (vla-get-ObjectName obj)))
        ;; Apply entity-type filter (nil onames = search everything)
        (if (or searchAll (member oname onames))
          (progn
            (setq txt (halff:vla-get-text obj))
            (if (and txt (halff:text-contains? textContent txt))
              (progn
                (setq found (1+ found))

                (setq hitCount 0 firstHitIdx -1 i 0)
                (while (< i (length vps))
                  (setq inVpRes
                    (vl-catch-all-apply 'halff:vla-ent-in-vp
                                        (list obj (nth i vps))))
                  (if (and (not (vl-catch-all-error-p inVpRes)) inVpRes)
                    (progn
                      (if (= firstHitIdx -1) (setq firstHitIdx i))
                      (setq hitCount (1+ hitCount))))
                  (setq i (1+ i)))

                (cond
                  ((= hitCount 1)
                   (setq totalQty (+ totalQty 1.0))
                   (setq vpQtys
                     (halff:setnth vpQtys (1+ firstHitIdx)
                                   (+ (nth firstHitIdx vpQtys) 1.0))))
                  ((> hitCount 1)
                   (setq totalQty (+ totalQty 1.0)))
                  (T nil)) ; outside all VPs — excluded from totalQty

                ;; Cache entity name for QHILITE / QSEARCH (current dwg only)
                (if isCurrentDwg
                  (progn
                    (setq en (vl-catch-all-apply
                               'vlax-vla-object->ename (list obj)))
                    (if (not (vl-catch-all-error-p en))
                      (halff:remember-entity payitem obj))))
                      ))))
                      )

      (princ (strcat "\n      Matched " (itoa found)
                     " text entities containing \""
                     textContent "\""))
      (if (and (not isCurrentDwg) dbx)
        (vl-catch-all-apply 'vlax-release-object (list dbx)))
      (list totalQty vpQtys))))

;; ===========================================
;; MAIN QTO COMMAND
;; ===========================================


;; ===========================================
;; TEXT QTY HELPERS
;; ===========================================

(defun halff:numeric-str? (s / i ch valid hasdot hasdigit)
  (setq s (halff:trim s))
  (if (= (strlen s) 0)
    nil
    (progn
      (setq valid T i 1 hasdot nil hasdigit nil)
      (while (and valid (<= i (strlen s)))
        (setq ch (ascii (substr s i 1)))
        (cond
          ((and (= i 1) (member ch (list 43 45)))
           nil)
          ((= ch 46)
           (if hasdot (setq valid nil) (setq hasdot T)))
          ((and (>= ch 48) (<= ch 57))
           (setq hasdigit T))
          (T (setq valid nil)))
        (setq i (1+ i)))
      (and valid hasdigit))))

(defun halff:extract-qty-from-text
       (pattern placeholder text
        / nwspat nwsph nwstxt phpos before after
          start-idx end-idx numstr)
  ;; Returns the number at the placeholder position in text, or nil.
  ;; pattern="PROP. # LF STEEL ENCASEMENT", placeholder="#",
  ;; text="PROP. 45 LF STEEL ENCASEMENT" -> 45.0
  (setq nwspat (strcase (halff:normalize-ws pattern))
        nwsph  (strcase (halff:normalize-ws placeholder))
        nwstxt (strcase (halff:normalize-ws text)))
  (setq phpos (vl-string-search nwsph nwspat))
  (if (null phpos)
    nil
    (progn
      (setq before (substr nwspat 1 phpos))
      (setq after  (substr nwspat (+ phpos (strlen nwsph) 1)))
      (if (= before "")
        (setq start-idx 0)
        (progn
          (setq start-idx (vl-string-search before nwstxt))
          (if start-idx
            (setq start-idx (+ start-idx (strlen before)))
            (setq start-idx nil))))
      (if (null start-idx)
        nil
        (progn
          (if (= after "")
            (setq end-idx (strlen nwstxt))
            (setq end-idx (vl-string-search after nwstxt start-idx)))
          (if (null end-idx)
            nil
            (progn
              (setq numstr (halff:trim
                             (substr nwstxt (1+ start-idx)
                                     (- end-idx start-idx))))
              (if (halff:numeric-str? numstr)
                (atof numstr)
                nil))))))))

;; -----------------------------------------------------------------------
(defun halff:sum-text-qty-in-file
       (dwgPath payitem pattern placeholder objStr vps
        / dbx openRes ms onames searchAll obj oname txt
          qtyVal found totalQty vpQtys hitCount firstHitIdx i
          inVpRes isCurrentDwg curPath en)
  (setq curPath     (strcat (getvar "DWGPREFIX") (getvar "DWGNAME"))
        isCurrentDwg (= (strcase dwgPath) (strcase curPath)))
  (if isCurrentDwg
    (progn
      (princ "\n    Processing CURRENT drawing (text-qty mode)")
      (setq ms (vla-get-ModelSpace
                 (vla-get-ActiveDocument (vlax-get-acad-object)))))
    (progn
      (setq dbx (halff:dbx-create))
      (if (not dbx)
        (progn
          (princ (strcat "\n    ERROR: Cannot create DBX for " dwgPath))
          (halff:log-failure payitem dwgPath "" objStr "N/A"
                             "Cannot create DBX object")
          (setq ms nil))
        (progn
          (setq openRes (vl-catch-all-apply 'vla-open (list dbx dwgPath)))
          (if (vl-catch-all-error-p openRes)
            (progn
              (princ (strcat "\n    ERROR: Cannot open " dwgPath))
              (halff:log-failure payitem dwgPath "" objStr "N/A"
                                 (strcat "Cannot open: "
                                         (vl-catch-all-error-message openRes)))
              (vl-catch-all-apply 'vlax-release-object (list dbx))
              (setq ms nil))
            (setq ms (vla-get-ModelSpace dbx)))))))
  (if (not ms)
    (list 0.0 (halff:zeros (length vps)))
    (progn
      (setq onames    (halff:text-obj-types objStr)
            searchAll (null onames))
      (princ (strcat "\n    Processing: " dwgPath))
      (princ (strcat "\n      Text-qty pattern: \"" pattern "\""))
      (princ (strcat "\n      Placeholder: \"" placeholder "\""))
      (princ (strcat "\n      Object filter: "
                     (if searchAll "ALL entities" objStr)))
      (setq found    0
            totalQty 0.0
            vpQtys   (halff:zeros (length vps)))
      (vlax-for obj ms
        (setq oname (strcase (vla-get-ObjectName obj)))
        (if (or searchAll (member oname onames))
          (progn
            (setq txt (halff:vla-get-text obj))
            (if txt
              (progn
                (setq qtyVal
                  (halff:extract-qty-from-text pattern placeholder txt))
                (if qtyVal
                  (progn
                    (setq found (1+ found))
                    (setq hitCount 0 firstHitIdx -1 i 0)
                    (while (< i (length vps))
                      (setq inVpRes
                        (vl-catch-all-apply 'halff:vla-ent-in-vp
                                            (list obj (nth i vps))))
                      (if (and (not (vl-catch-all-error-p inVpRes)) inVpRes)
                        (progn
                          (if (= firstHitIdx -1) (setq firstHitIdx i))
                          (setq hitCount (1+ hitCount))))
                      (setq i (1+ i)))
                    (cond
                      ((= hitCount 1)
                       (setq totalQty (+ totalQty qtyVal))
                       (setq vpQtys
                         (halff:setnth vpQtys (1+ firstHitIdx)
                                       (+ (nth firstHitIdx vpQtys) qtyVal))))
                      ((> hitCount 1)
                       (setq totalQty (+ totalQty qtyVal)))
                      (T nil)) ; outside all VPs — excluded
                    (if isCurrentDwg
                      (progn
                        (setq en (vl-catch-all-apply
                                   'vlax-vla-object->ename (list obj)))
                        (if (not (vl-catch-all-error-p en))
                          (halff:remember-entity payitem obj)))))))))))
      (princ (strcat "\n      Matched " (itoa found)
                     " entities, total qty=" (rtos totalQty 2 4)))
      (if (and (not isCurrentDwg) dbx)
        (vl-catch-all-apply 'vlax-release-object (list dbx)))
      (list totalQty vpQtys))))


;; ===========================================
;; CIVIL3D PIPE / STRUCTURE HELPERS
;; ===========================================

(defun halff:civil3d-qty-in-file
       (dwgPath payitem unit styleFilter descFilter objType vps
        / dbx openRes ms obj oname styleObj styleVal descVal descParts
          lenVal qtyVal found totalQty vpQtys hitCount firstHitIdx i
          inVpRes isCurrentDwg curPath en fs searchToken
          styleMatch descMatch
          vpNamesCrossing pipeSP pipeEP pipePts pipeIsFail
          res1 res2 px py mnV mxV mnL mxL)
  ;; objType: "PIPE" or "STRUCTURE"
  ;; styleFilter: Civil3D StyleName to match (case-insensitive)
  ;; unit: "LF" for pipes (sum Length2D), "EA" for structures (count)
  (setq curPath     (strcat (getvar "DWGPREFIX") (getvar "DWGNAME"))
        isCurrentDwg (= (strcase dwgPath) (strcase curPath))
        searchToken  (if (= (strcase objType) "PIPE") "pipe" "struct")
        fs           (halff:feet-scale))
  (if isCurrentDwg
    (progn
      (princ (strcat "\n    Processing CURRENT drawing (civil3d-" objType " mode)"))
      (setq ms (vla-get-ModelSpace
                 (vla-get-ActiveDocument (vlax-get-acad-object)))))
    (progn
      (setq dbx (halff:dbx-create))
      (if (not dbx)
        (progn
          (princ (strcat "\n    ERROR: Cannot create DBX for " dwgPath))
          (halff:log-failure payitem dwgPath "" objType "N/A"
                             "Cannot create DBX object")
          (setq ms nil))
        (progn
          (setq openRes (vl-catch-all-apply 'vla-open (list dbx dwgPath)))
          (if (vl-catch-all-error-p openRes)
            (progn
              (princ (strcat "\n    ERROR: Cannot open " dwgPath))
              (halff:log-failure payitem dwgPath "" objType "N/A"
                                 (strcat "Cannot open: "
                                         (vl-catch-all-error-message openRes)))
              (vl-catch-all-apply 'vlax-release-object (list dbx))
              (setq ms nil))
            (setq ms (vla-get-ModelSpace dbx)))))))
  (if (not ms)
    (list 0.0 (halff:zeros (length vps)))
    (progn
      (princ (strcat "\n    Processing: " dwgPath))
      (princ (strcat "\n      Civil3D " objType
                     " style: \"" styleFilter "\""
                     (if (/= descFilter "")
                       (strcat "  desc: \"" descFilter "\"")
                       "")))
      (setq found    0
            totalQty 0.0
            vpQtys   (halff:zeros (length vps)))
      (vlax-for obj ms
        ;; Match by VLA ObjectName OR DXF entity type (handles proxy entities)
        (setq oname (vl-catch-all-apply 'vla-get-ObjectName (list obj)))
        (setq en    (vl-catch-all-apply 'vlax-vla-object->ename (list obj)))
        (setq dxftype
          (if (not (vl-catch-all-error-p en))
            (strcase (cdr (assoc 0 (entget en))))
            ""))
        (if (or (and (not (vl-catch-all-error-p oname))
                     (vl-string-search searchToken (strcase oname T)))
                (vl-string-search searchToken dxftype))
          (progn
            ;; --- Style match ------------------------------------------
            (setq styleObj (vl-catch-all-apply
                             'vlax-get-property (list obj "Style")))
            (if (not (vl-catch-all-error-p styleObj))
              (setq styleVal (vl-catch-all-apply
                               'vlax-get-property (list styleObj "Name")))
              (setq styleVal nil))
            (setq styleMatch
              (if (= styleFilter "")
                T  ; no style filter -> pass
                (and styleVal (not (vl-catch-all-error-p styleVal))
                     (= (type styleVal) 'STR)
                     (vl-some
                       '(lambda (tok)
                          (= (strcase (halff:trim (halff:normalize-ws tok)))
                             (strcase (halff:trim (halff:normalize-ws styleVal)))))
                       (halff:split styleFilter ",")))))
            ;; --- Description match ------------------------------------
            (setq descVal (vl-catch-all-apply
                            'vlax-get-property (list obj "Description")))
            (if (or (vl-catch-all-error-p descVal) (null descVal))
              (setq descVal ""))
            (setq descMatch
              (if (= descFilter "")
                T  ; no desc filter -> pass
                (progn
                  ;; Split descFilter on commas; match if any token equals
                  ;; the object's Description (case-insensitive, trimmed)
                  (setq descParts (halff:split descFilter ","))
                  (setq descVal (halff:normalize-ws (vl-princ-to-string descVal)))
                  (vl-some
                    '(lambda (tok)
                       (= (strcase (halff:trim (halff:normalize-ws tok)))
                          (strcase (halff:trim descVal))))
                    descParts))))
            (if (and styleMatch descMatch)
              (progn
                (if (and (= (strcase objType) "PIPE")
                         (= (strcase (halff:trim unit)) "LF"))
                  (progn
                    ;; LF pipe: sum Length2D (fall back to Length3D then Length)
                    (setq lenVal (vl-catch-all-apply
                                   'vlax-get-property (list obj "Length2D")))
                    (if (or (vl-catch-all-error-p lenVal) (null lenVal))
                      (setq lenVal (vl-catch-all-apply
                                     'vlax-get-property (list obj "Length3D"))))
                    (if (or (vl-catch-all-error-p lenVal) (null lenVal))
                      (setq lenVal (vl-catch-all-apply
                                     'vlax-get-property (list obj "Length"))))
                    (if (and (not (vl-catch-all-error-p lenVal))
                             lenVal
                             (member (type lenVal) '(REAL INT)))
                      (setq qtyVal (* (float lenVal) fs))
                      (setq qtyVal 0.0))
                    (if *HALFF_ROUND_EACH*
                      (setq qtyVal (halff:round0 qtyVal))))
                  ;; EA (pipe count, structure count, or any other unit): count
                  (setq qtyVal 1.0))
                (setq found (1+ found))
                ;; VP attribution: only count objects inside exactly 1 viewport
                (setq hitCount 0 firstHitIdx -1 i 0)
                (while (< i (length vps))
                  (setq inVpRes
                    (vl-catch-all-apply 'halff:vla-ent-in-vp
                                        (list obj (nth i vps))))
                  (if (and (not (vl-catch-all-error-p inVpRes)) inVpRes)
                    (progn
                      (if (= firstHitIdx -1) (setq firstHitIdx i))
                      (setq hitCount (1+ hitCount))))
                  (setq i (1+ i)))
                ;; Crossing detection for pipes: get start/end as 2D pts,
                ;; then check if the segment crosses any VP boundary polygon.
                ;; Three fallback methods to handle Civil3D proxy variance.
                (setq vpNamesCrossing '() pipeIsFail nil pipePts nil)
                (if (= (strcase objType) "PIPE")
                  (progn
                    (setq pipeSP nil pipeEP nil)
                    ;; Method 1: vlax-curve (standard AcDbCurve interface)
                    (if (not (vl-catch-all-error-p en))
                      (progn
                        (setq res1 (vl-catch-all-apply 'vlax-curve-getStartPoint (list en)))
                        (setq res2 (vl-catch-all-apply 'vlax-curve-getEndPoint (list en)))
                        (if (and (not (vl-catch-all-error-p res1)) (listp res1))
                          (setq pipeSP (list (car res1) (cadr res1))))
                        (if (and (not (vl-catch-all-error-p res2)) (listp res2))
                          (setq pipeEP (list (car res2) (cadr res2))))))
                    ;; Method 2: vlax-get-property — SafeArray, VARIANT, or VLA-OBJECT
                    (if (not pipeSP)
                      (progn
                        (setq res1 (vl-catch-all-apply 'vlax-get-property
                                                        (list obj "StartPoint")))
                        (setq pipeSP
                          (cond
                            ((or (null res1) (vl-catch-all-error-p res1)) nil)
                            ((listp res1) (list (car res1) (cadr res1)))
                            ((member (type res1) '(VARIANT SAFEARRAY))
                             (setq res1 (halff:variant->list res1))
                             (if (listp res1) (list (car res1) (cadr res1)) nil))
                            ((= (type res1) 'VLA-OBJECT)
                             (setq px (vl-catch-all-apply 'vlax-get-property
                                                           (list res1 "X")))
                             (setq py (vl-catch-all-apply 'vlax-get-property
                                                           (list res1 "Y")))
                             (if (and (not (vl-catch-all-error-p px)) (numberp px)
                                      (not (vl-catch-all-error-p py)) (numberp py))
                               (list px py) nil))
                            (T nil)))))
                    (if (not pipeEP)
                      (progn
                        (setq res2 (vl-catch-all-apply 'vlax-get-property
                                                        (list obj "EndPoint")))
                        (setq pipeEP
                          (cond
                            ((or (null res2) (vl-catch-all-error-p res2)) nil)
                            ((listp res2) (list (car res2) (cadr res2)))
                            ((member (type res2) '(VARIANT SAFEARRAY))
                             (setq res2 (halff:variant->list res2))
                             (if (listp res2) (list (car res2) (cadr res2)) nil))
                            ((= (type res2) 'VLA-OBJECT)
                             (setq px (vl-catch-all-apply 'vlax-get-property
                                                           (list res2 "X")))
                             (setq py (vl-catch-all-apply 'vlax-get-property
                                                           (list res2 "Y")))
                             (if (and (not (vl-catch-all-error-p px)) (numberp px)
                                      (not (vl-catch-all-error-p py)) (numberp py))
                               (list px py) nil))
                            (T nil)))))
                    ;; Method 3: bbox diagonal as last resort
                    (if (or (not pipeSP) (not pipeEP))
                      (progn
                        (setq res1 (vl-catch-all-apply
                                     '(lambda ()
                                        (vla-getboundingbox obj 'mnV 'mxV)
                                        (list mnV mxV))
                                     nil))
                        (if (and (not (vl-catch-all-error-p res1)) res1)
                          (progn
                            (setq mnL (halff:variant->list mnV))
                            (setq mxL (halff:variant->list mxV))
                            (if (not pipeSP) (setq pipeSP (list (car mnL) (cadr mnL))))
                            (if (not pipeEP) (setq pipeEP (list (car mxL) (cadr mxL))))))))
                    (if (and pipeSP pipeEP)
                      (progn
                        (setq pipePts (list pipeSP pipeEP))
                        (setq i 0)
                        (while (< i (length vps))
                          (if (and (/= i firstHitIdx)
                                   (halff:genuine-crossing? pipePts (cadr (nth i vps))))
                            (setq vpNamesCrossing
                              (append vpNamesCrossing (list (car (nth i vps))))))
                          (setq i (1+ i))))
                      (princ "\n      NOTE: Could not get pipe endpoints for crossing check"))))
                (cond
                  ;; Pipe crosses a VP boundary — log failure, exclude from total
                  ((and (= (strcase objType) "PIPE") vpNamesCrossing)
                   (setq pipeIsFail T)
                   (princ (strcat "\n      WARNING: PIPE crosses viewport boundary"))
                   (halff:log-failure payitem dwgPath "" "PIPE"
                                      vpNamesCrossing "Pipe crosses viewport boundary"))
                  ((= hitCount 1)
                   (setq totalQty (+ totalQty qtyVal))
                   (setq vpQtys
                     (halff:setnth vpQtys (1+ firstHitIdx)
                                   (+ (nth firstHitIdx vpQtys) qtyVal))))
                  ((> hitCount 1)
                   (setq pipeIsFail T)
                   (princ (strcat "\n      WARNING: " objType
                                  " in " (itoa hitCount)
                                  " viewports - excluding from total"))
                   (if (= (strcase objType) "PIPE")
                     (halff:log-failure payitem dwgPath "" "PIPE"
                                        (itoa hitCount) "Pipe in multiple viewports")))
                  (T
                   (setq pipeIsFail T)
                   (princ (strcat "\n      WARNING: " objType
                                  " outside all viewports - excluded"))
                   (if (= (strcase objType) "PIPE")
                     (halff:log-failure payitem dwgPath "" "PIPE"
                                        "OUTSIDE" "Pipe outside all viewports"))))
                (if (and isCurrentDwg (= hitCount 1))
                  (progn
                    (setq en (vl-catch-all-apply
                               'vlax-vla-object->ename (list obj)))
                    (if (not (vl-catch-all-error-p en))
                      (halff:remember-entity payitem obj)))))))))
      (princ (strcat "\n      Matched " (itoa found)
                     " " objType "(s), total=" (rtos totalQty 2 4)))
      (if (and (not isCurrentDwg) dbx)
        (vl-catch-all-apply 'vlax-release-object (list dbx)))
      (list totalQty vpQtys))))

(defun halff:qto-to-csv
       (xlsx
        / outcsv outfail xl wb ws maxc lastRow hdr
          cPay cDesc cUnit cLayer cObj cPattern cLtype
          cMult cTextContent cTextQty cStyle cC3dDesc cFormula cFilePath
          r idx rowvals pay desc unit layer obj pattern ltype
          mult textcontent textqty style c3ddesc formula filepath
          processedRows rawMap pr rawQty rawVpQtys finalQty finalVpQtys fval
          vps result totalQty vpVals outrow fs f xlPath xlDir xlBase)

  (princ "\n+===============================================================+")
  (princ "\n|           HALFF QTO LABELS v4.0 - MULTI-FILE                  |")
  (princ "\n+===============================================================+")
  (if (not *HALFF_VP_DEFS*)
    (progn
      (princ "\n[QRUN] No viewports loaded. Loading from sidecar...")
      (c:QVPLOAD)))
  (if (not *HALFF_VP_DEFS*)
    (progn
      (princ "\n[QRUN] ERROR: No viewports loaded. Run QVPLOAD first.")
      nil)
    (progn
      (setq vps (halff:vp-prep *HALFF_VP_DEFS*))
      (princ (strcat "\n[QRUN] Loaded " (itoa (length vps)) " viewports"))
      (setq xlPath xlsx)
      (setq xlDir (vl-filename-directory xlPath))
      (if (not xlDir) (setq xlDir (getvar "DWGPREFIX")))
      (setq xlBase (vl-filename-base xlPath))
      (setq outcsv  (strcat xlDir "\\" xlBase "_QTO.csv"))
      (setq outfail (strcat xlDir "\\" xlBase "_QTO_FAILURES.csv"))
      (setq *HALFF_LAST_CSV* outcsv)
      (setq *HALFF_FAILURE_LOG* nil)
      (setq *HALFF_PAYITEM_ENTS* nil)
      (setq *HALFF_FAIL_ENTS* nil)
      (setq fs (halff:feet-scale))
      (princ (strcat "\n[QRUN] Reading mapping: " xlsx))
      (princ (strcat "\n[QRUN] Output CSV: " outcsv))
      (princ (strcat "\n[QRUN] Failure CSV: " outfail))
      (princ (strcat "\n[QRUN] Units: INSUNITS=" (itoa (getvar "INSUNITS"))
                     " feetScale=" (rtos fs 2 4)
                     " roundEach=" (if *HALFF_ROUND_EACH* "YES" "NO")))
      (if (findfile outcsv)  (vl-catch-all-apply 'vl-file-delete (list outcsv)))
      (if (findfile outfail) (vl-catch-all-apply 'vl-file-delete (list outfail)))
      (setq xlwb (vl-catch-all-apply 'halff:xl-open-ro (list xlsx)))
      (if (vl-catch-all-error-p xlwb)
        (progn
          (princ (strcat "\nQRUN ERROR (Excel.Open): "
                         (vl-catch-all-error-message xlwb)))
          nil)
        (progn
          (setq xl (car xlwb) wb (cadr xlwb))
          (setq ws (halff:ws-active wb))
          (setq maxc    (halff:used-cols ws))
          (setq lastRow (halff:used-rows ws))
          (princ (strcat "\n[QRUN] UsedRange rows=" (itoa lastRow)
                         " cols=" (itoa maxc)))
          (princ "\n[QRUN] Reading headers...")
          (setq hdr (vl-catch-all-apply 'halff:headers (list ws)))
          (if (vl-catch-all-error-p hdr)
            (progn
              (princ (strcat "\n[QRUN] ERROR reading headers: "
                             (vl-catch-all-error-message hdr)))
              (halff:xl-close xl wb)
              nil)
            (progn
              (princ (strcat "\n[QRUN] Headers read: "
                             (itoa (length hdr)) " columns"))

              ;; -- Column index lookup --------------------------------
              (setq cPay         (halff:col hdr "PAY ITEM"))
              (setq cDesc        (halff:col hdr "ITEM DESCRIPTION"))
              (if (not cDesc) (setq cDesc (halff:col hdr "ITEM DES")))
              (setq cUnit        (halff:col hdr "UNIT"))
              (setq cLayer       (halff:col hdr "LAYER"))
              (setq cObj         (halff:col hdr "OBJECT"))
              (setq cPattern     (halff:col hdr "PATTERN NAME"))
              (setq cLtype       (halff:col hdr "LINETYPE"))
              (setq cMult        (halff:col hdr "MULTIPLIER"))
              (setq cTextContent (halff:col hdr "TEXT CONTENT"))
              (setq cTextQty     (halff:col hdr "TEXT QTY"))
              (setq cStyle      (halff:col hdr "PIPE OR STRUCTURE STYLE"))
              (setq cC3dDesc    (halff:col hdr "DESCRIPTION"))
              (setq cFormula    (halff:col hdr "FORMULA"))
              (setq cFilePath    (halff:col hdr "FILE PATH"))

              ;; PAY ITEM, UNIT, OBJECT, FILE PATH are always required.
              ;; LAYER is required per-row only for geometry-mode rows.
              ;; TEXT CONTENT is optional at the header level.
              (if (not (and cPay cUnit cObj cFilePath))
                (progn
                  (princ "\nMissing required headers. Need: PAY ITEM, UNIT, OBJECT, FILE PATH.")
                  (halff:xl-close xl wb)
                  nil)
                (progn
                  (setq f (open outcsv "w"))
                  (if (not f)
                    (progn
                      (princ (strcat "\nERROR: Could not write CSV: " outcsv))
                      (halff:xl-close xl wb)
                      nil)
                    (progn
                      ;; Write CSV header row
                      (setq outrow (list "PAY ITEM" "ITEM DESCRIPTION"
                                         "UNIT" "QTY_MODEL"))
                      (foreach vp *HALFF_VP_DEFS*
                        (setq outrow (append outrow
                                             (list (strcat "QTY_" (car vp))))
                                             ))
                      (write-line (halff:csv-line outrow) f)
                      (setq processedRows '())
                      (princ "\n[QRUN] Processing rows...")
                      (setq r 2)
                      (while (<= r lastRow)
                        ;; Read all cells in this row
                        (setq rowvals '() idx 1)
                        (while (<= idx maxc)
                          (setq rowvals
                            (append rowvals
                                    (list (halff:get-row-string ws r idx))))
                          (setq idx (1+ idx)))

                        ;; Extract column values
                        (setq pay
                          (halff:trim (nth (1- cPay) rowvals)))
                        (setq desc
                          (if cDesc (halff:trim (nth (1- cDesc) rowvals)) ""))
                        (setq unit
                          (halff:trim (nth (1- cUnit) rowvals)))
                        (setq layer
                          (if cLayer (halff:trim (nth (1- cLayer) rowvals)) ""))
                        (setq obj
                          (halff:trim (nth (1- cObj) rowvals)))
                        (setq pattern
                          (if cPattern (halff:trim (nth (1- cPattern) rowvals)) ""))
                        (setq ltype
                          (if cLtype (halff:trim (nth (1- cLtype) rowvals)) ""))
                        (setq mult
                          (if cMult (halff:trim (nth (1- cMult) rowvals)) ""))
                        (setq textcontent
                          (if cTextContent
                            (halff:trim (nth (1- cTextContent) rowvals))
                            ""))
                        (setq textqty
                          (if cTextQty
                            (halff:trim (nth (1- cTextQty) rowvals))
                            ""))
                        (setq style
                          (if cStyle
                            (halff:trim (nth (1- cStyle) rowvals))
                            ""))
                        (setq c3ddesc
                          (if cC3dDesc
                            (halff:trim (nth (1- cC3dDesc) rowvals))
                            ""))
                        (setq formula
                          (if cFormula
                            (halff:trim (nth (1- cFormula) rowvals))
                            ""))
                        (setq filepath
                          (halff:trim (nth (1- cFilePath) rowvals)))

                        ;; A row is valid when it has the core fields AND
                        ;; either a layer (geometry mode) or text content
                        ;; (text-count mode).
                        (if (and (/= pay "")
                                 (/= unit "")
                                 (/= obj "")
                                 (/= filepath "")
                                 (or (/= layer "")
                                     (/= textcontent "")
                                     (and (or (/= style "") (/= c3ddesc ""))
                                          (member (strcase obj)
                                                  (list "PIPE" "STRUCTURE")))))
                          (progn
                            (princ (strcat "\n[ROW " (itoa r)
                                           "] PAY=" pay
                                           " UNIT=" unit
                                           (if (/= textcontent "")
                                             (strcat " TEXT=\"" textcontent "\"")
                                             (strcat " LAYER=" layer))
                                           " FILE=" filepath))
                            (if (not (findfile filepath))
                              (progn
                                (princ (strcat "\n    WARNING: File not found: "
                                               filepath))
                                (halff:log-failure pay filepath layer obj
                                                   "N/A" "File not found")
                                (setq result
                                  (list 0.0 (halff:zeros (length vps))))
                                  )
                              ;; --- Route to correct processor --------
                              (cond
                                ((member (strcase obj) (list "PIPE" "STRUCTURE"))
                                 (setq result
                                   (halff:civil3d-qty-in-file
                                     filepath pay unit style c3ddesc obj vps)))
                                ((and (/= textcontent "") (/= textqty ""))
                                 (setq result
                                   (halff:sum-text-qty-in-file
                                     filepath pay textcontent textqty
                                     obj vps)))
                                ((/= textcontent "")
                                 (setq result
                                   (halff:count-text-in-file
                                     filepath pay textcontent obj vps)))
                                (T
                                 (setq result
                                   (halff:process-dwg-file
                                     filepath pay unit layer obj
                                     ltype pattern mult vps)))))
                            ;; Store result for pass 2
                            (setq processedRows
                              (append processedRows
                                      (list (list r pay desc unit formula
                                                  (car result) (cadr result)))))))
                        (setq r (1+ r)))

                      ;; --- Pass 2: apply formulas and write output --------
                      ;; Build raw qty map: (paystr rawModelQty)
                      (setq rawMap
                        (mapcar '(lambda (pr) (list (nth 1 pr) (nth 5 pr)))
                                processedRows))
                      (foreach pr processedRows
                        (setq r         (nth 0 pr)
                              pay       (nth 1 pr)
                              desc      (nth 2 pr)
                              unit      (nth 3 pr)
                              formula   (nth 4 pr)
                              rawQty    (nth 5 pr)
                              rawVpQtys (nth 6 pr))
                        (if (/= formula "")
                          (progn
                            (setq fval (halff:formula-eval formula rawQty rawMap))
                            (setq finalQty    (if fval (fix (+ fval 0.5)) rawQty))
                            (setq finalVpQtys
                              (mapcar '(lambda (vq / fv)
                                         (setq fv (halff:formula-eval
                                                    formula vq rawMap))
                                         (if fv (fix (+ fv 0.5)) vq))
                                      rawVpQtys))
                            (princ (strcat "\n[ROW " (itoa r) "] FORMULA "
                                           formula " → x=" (rtos rawQty 2 4)
                                           " result=" (rtos finalQty 2 4))))
                          (progn
                            (setq finalQty    rawQty)
                            (setq finalVpQtys rawVpQtys)))
                        (setq outrow (list pay desc unit (rtos finalQty 2 4)))
                        (foreach vpq finalVpQtys
                          (setq outrow (append outrow (list (rtos vpq 2 4)))))
                        (write-line (halff:csv-line outrow) f))

                      (close f)
                      (halff:xl-close xl wb)
                      (if *HALFF_FAILURE_LOG*
                        (progn
                          (setq f (open outfail "w"))
                          (write-line
                            "PAY ITEM,FILE PATH,LAYER,OBJECT TYPE,VIEWPORT,REASON"
                            f)
                          (foreach fail *HALFF_FAILURE_LOG*
                            (write-line (halff:csv-line fail) f))
                          (close f)
                          (princ (strcat "\n[QRUN] OK Failure log: " outfail))
                          (princ (strcat "\n[QRUN] OK Total failures: "
                                         (itoa (length *HALFF_FAILURE_LOG*))))
                                         ))
                      (princ "\n+===============================================================+")
                      (princ "\n|                   QTO COMPLETE!                               |")
                      (princ "\n+===============================================================+")
                      (princ (strcat "\n[QRUN] OK CSV written: " outcsv))
                      (princ (strcat "\n[QRUN] OK Rows processed: "
                                     (itoa (- lastRow 1))))
                      T))))
                      ))))
                      )))

;; ===========================================
;; USER COMMANDS
;; ===========================================

(defun c:QMAPSET (/ cur p fp ok)
  (setq fp (halff:sidecar-path))
  (setq cur (halff:get-mapping-path))
  (if (and cur (/= cur "")) (princ (strcat "\nCurrent mapping (sidecar): " cur)))
  (setq p (getfiled "Select QTO mapping Excel (.xlsx)" (if cur cur "") "xlsx" 0))
  (if (and p (/= p ""))
    (progn
      (setq ok (halff:set-mapping-path p))
      (if ok (progn (princ (strcat "\nMapping set: " p)) (princ (strcat "\nSidecar file: " fp)))
        (princ (strcat "\nERROR: Could not write sidecar file: " fp))))
    (princ "\nNo file selected."))
  (princ))

(defun c:QRUN (/ p)
  (setq p (halff:get-mapping-path))
  (if (and p (/= p "")) (princ (strcat "\nUsing mapping (sidecar): " p)))
  (if (and p (/= p "") (findfile p))
    (halff:qto-to-csv p)
    (princ "\nNo mapping set or file not found. Run QMAPSET."))
  (princ))

(defun c:QOPENREPORT (/ p ok)
  (setq p *HALFF_LAST_CSV*)
  (if (and p (findfile p))
    (progn
      (princ (strcat "\nOpening: " p))
      (startapp "explorer.exe" (strcat "/select,\"" p "\"")))
    (princ "\nNo CSV found to open yet."))
  (princ))

(defun c:QVPCROSSZERO (/)
  (setq *HALFF_CROSSING_QTY_ZERO* T)
  (princ "\nOK Crossing entities will be ZEROED OUT (excluded from quantities)")
  (princ))

(defun c:QVPCROSSCOUNT (/)
  (setq *HALFF_CROSSING_QTY_ZERO* nil)
  (princ "\nOK Crossing entities will be COUNTED in their assigned viewport")
  (princ))

(defun c:QVPCROSS (/)
  (princ "\n+========================================================+")
  (princ "\n|  CROSSING DETECTION SETTINGS                           |")
  (princ "\n+========================================================+")
  (princ (strcat "\n|  Mode: " (if *HALFF_CROSSING_QTY_ZERO* "ZERO OUT" "COUNT   ")
                "                                      |"))
  (princ (strcat "\n|  Exclusion: " (rtos *HALFF_EXCLUSION_DIST* 2 3) " feet                              |"))
  (princ "\n+========================================================+")
  (princ "\n|  -QVPCROSSZERO   - Zero out crossing entities           |")
  (princ "\n|  -QVPCROSSCOUNT  - Count crossing entities              |")
  (princ "\n|  -QVPEXCLUDE     - Set exclusion distance              |")
  (princ "\n+========================================================+")
  (princ))

(defun c:QVPEXCLUDE (/ dist)
  (initget 6)
  (setq dist (getdist (strcat "\nEnter exclusion distance in feet <"
                              (rtos *HALFF_EXCLUSION_DIST* 2 3) ">: ")))
  (if dist
    (progn
      (setq *HALFF_EXCLUSION_DIST* dist)
      (princ (strcat "\nOK Exclusion distance set to "
                     (rtos *HALFF_EXCLUSION_DIST* 2 3) " feet")))
    (princ (strcat "\nOK Current exclusion distance: "
                   (rtos *HALFF_EXCLUSION_DIST* 2 3) " feet")))
  (princ))

;; ===========================================
;; HIGHLIGHTING & SEARCH COMMANDS
;; ===========================================

(defun halff:remember-entity (payitem vlaObj / en rec lst)
  (if (= (type payitem) 'STR)
    (progn
      (setq payitem (vl-string-trim " " payitem))
      (if (wcmatch payitem "*#*")
        (setq payitem (rtos (atof payitem) 2 1))))
        )
  (if (and *HALFF_HIGHLIGHT_ENABLE* payitem (/= payitem ""))
    (progn
      (setq en (vl-catch-all-apply 'vlax-vla-object->ename (list vlaObj)))
      (if (not (vl-catch-all-error-p en))
        (progn
          (if (and en (entget en))
            (progn
              (setq rec (assoc payitem *HALFF_PAYITEM_ENTS*))
              (if rec
                (progn
                  (setq lst (cdr rec))
                  (if (not (member en lst))
                    (setq *HALFF_PAYITEM_ENTS*
                          (subst (cons payitem (append lst (list en)))
                                 rec
                                 *HALFF_PAYITEM_ENTS*))))
                (setq *HALFF_PAYITEM_ENTS*
                      (cons (cons payitem (list en)) *HALFF_PAYITEM_ENTS*))))
                      ))))
                      )
  nil)

(defun halff:remember-fail-entity (vlaObj / en)
  (setq en (vl-catch-all-apply 'vlax-vla-object->ename (list vlaObj)))
  (if (not (vl-catch-all-error-p en))
    (if (and en (entget en) (not (member en *HALFF_FAIL_ENTS*)))
      (setq *HALFF_FAIL_ENTS* (append *HALFF_FAIL_ENTS* (list en))))
      )
  nil)

(defun halff:payitem->ss (pay / rec ss en)
  (setq rec (assoc pay *HALFF_PAYITEM_ENTS*))
  (if rec
    (progn
      (setq ss (ssadd))
      (foreach en (cdr rec)
        (if (and en (entget en))
          (ssadd en ss)))
      ss)
    nil))

(defun halff:resolve-paykey (s / s0 d num fmt)
  (setq s0 (vl-string-trim " " s))
  (setq d (vl-string-search "." s0))
  (if d
    (progn
      (setq num (atof s0))
      (setq fmt (rtos num 2 1))
      (if (assoc fmt *HALFF_PAYITEM_ENTS*)
        fmt
        (if (assoc s0 *HALFF_PAYITEM_ENTS*) s0 nil)))
    (if (assoc s0 *HALFF_PAYITEM_ENTS*)
      s0
      (progn
        (setq num (atof s0))
        (setq fmt (rtos num 2 1))
        (if (assoc fmt *HALFF_PAYITEM_ENTS*) fmt nil))))
        )

(defun halff:any->str (x)
  (if x (if (= (type x) 'STR) x (vl-princ-to-string x)) ""))

(defun halff:zoom-to-ename (en / obj mn mx mnL mxL dx dy pad p1 p2 acadObj)
  (setq obj (vlax-ename->vla-object en))
  (vla-getboundingbox obj 'mn 'mx)
  (setq mnL (halff:variant->list mn))
  (setq mxL (halff:variant->list mx))
  (setq dx (- (car mxL) (car mnL)))
  (setq dy (- (cadr mxL) (cadr mnL)))
  (setq pad (* 0.5 (max dx dy)))
  (setq p1 (list (- (car mnL) pad) (- (cadr mnL) pad) 0.0))
  (setq p2 (list (+ (car mxL) pad) (+ (cadr mxL) pad) 0.0))
  (setq acadObj (vlax-get-acad-object))
  (vla-ZoomWindow acadObj (vlax-3d-point p1) (vlax-3d-point p2)))

(defun halff:tol->prec (tol / s p dotpos)
  (setq s (rtos tol 2 8))
  (setq dotpos (vl-string-search "." s))
  (if dotpos
    (setq p (- (strlen s) dotpos 1))
    (setq p 0))
  (max p 0))

(defun halff:dup-key (e prec / dxf typ lay obj mn mx mnL mxL res)
  (setq dxf (entget e))
  (setq typ (cdr (assoc 0 dxf)))
  (setq lay (cdr (assoc 8 dxf)))
  (setq obj (vlax-ename->vla-object e))
  (setq res (vl-catch-all-apply 'vla-getboundingbox (list obj 'mn 'mx)))
  (if (vl-catch-all-error-p res)
    (strcat typ "|" lay "|" (vla-get-Handle obj))
    (progn
      (setq mnL (halff:variant->list mn))
      (setq mxL (halff:variant->list mx))
      (strcat typ "|" lay "|"
              (rtos (car mnL) 2 prec) ","
              (rtos (cadr mnL) 2 prec) ","
              (rtos (car mxL) 2 prec) ","
              (rtos (cadr mxL) 2 prec))))
              )

(defun halff:vp-name-list (/ out rec)
  (setq out '())
  (foreach rec *HALFF_VP_DEFS*
    (setq out (append out (list (car rec))))
    )
  out)

(defun halff:get-vp-poly-by-name (vpName / rec)
  (setq rec (assoc vpName *HALFF_VP_DEFS*))
  (if rec (cdr rec) nil))

(defun halff:payitem->ss-vp (pay vpPoly / rec ss en obj)
  (setq rec (assoc pay *HALFF_PAYITEM_ENTS*))
  (if rec
    (progn
      (setq ss (ssadd))
      (foreach en (cdr rec)
        (if (and en (entget en))
          (progn
            (setq obj (vlax-ename->vla-object en))
            (if (halff:vla-in-poly obj vpPoly)
              (ssadd en ss))))
              )
      ss)
    nil))

(defun halff:vla-in-poly (vlaObj poly / bb mn mx mnL mxList cx cy)
  (vla-getboundingbox vlaObj 'mn 'mx)
  (setq mnL (halff:variant->list mn))
  (setq mxList (halff:variant->list mx))
  (setq cx (/ (+ (car mnL) (car mxList)) 2.0))
  (setq cy (/ (+ (cadr mnL) (cadr mxList)) 2.0))
  (halff:pt-in-poly (list cx cy) poly))

(defun halff:join (lst sep / out)
  (setq out "")
  (foreach x lst
    (if (= out "")
      (setq out x)
      (setq out (strcat out sep x))))
  out)

;; ===========================================
;; USER COMMANDS - HIGHLIGHTING
;; ===========================================

(defun c:QHILITE (/ payRaw payKey ss prompt)
  (if (not *HALFF_PAYITEM_ENTS*)
    (princ "\nX No highlight data found. Run QRUN first.")
    (progn
      (princ (strcat "\nAvailable pay items cached: "
                     (itoa (length *HALFF_PAYITEM_ENTS*))
                     " (type exactly as in mapping)"))
      (setq prompt (strcat "\nEnter PAY ITEM to highlight"
                           (if *HALFF_LAST_HILITE_PAY*
                             (strcat " <"
                                     (halff:any->str *HALFF_LAST_HILITE_PAY*)
                                     ">: ")
                             ": ")))
      (setq payRaw (getstring T prompt))
      (if (or (not payRaw) (= payRaw ""))
        (setq payRaw (if *HALFF_LAST_HILITE_PAY*
                       (halff:any->str *HALFF_LAST_HILITE_PAY*) "")))
      (if (or (not payRaw) (= payRaw ""))
        (princ "\nX No pay item entered.")
        (progn
          (setq payKey (halff:resolve-paykey payRaw))
          (if (not payKey) (setq payKey payRaw))
          (setq ss (halff:payitem->ss payKey))
          (if (and ss (> (sslength ss) 0))
            (progn
              (setq *HALFF_LAST_HILITE_PAY* payKey)
              (sssetfirst nil ss)
              (princ (strcat "\nOK Highlighted " (itoa (sslength ss))
                             " entities for PAY ITEM: "
                             (halff:any->str payKey))))
            (princ (strcat "\nX Pay item not found in this drawing: "
                           (halff:any->str payKey))))
                           ))))
  (princ))

(defun c:QHILITEVP (/ vpName vpPoly payRaw payKey ss vpNames)
  (vl-load-com)
  (if (not *HALFF_PAYITEM_ENTS*)
    (princ "\nX No highlight data found. Run QRUN first.")
    (if (not *HALFF_VP_DEFS*)
      (princ "\nX No viewport definitions loaded. Run QVPLOAD first.")
      (progn
        (setq vpNames (halff:vp-name-list))
        (princ (strcat "\nAvailable VPs: " (halff:join vpNames ", ")))
        (setq vpName (getstring T
                       (strcat "\nEnter VP name"
                               (if *HALFF_LAST_HILITE_VP*
                                 (strcat " <" *HALFF_LAST_HILITE_VP* ">: ")
                                 ": "))))
        (if (or (not vpName) (= (vl-string-trim " " vpName) ""))
          (setq vpName *HALFF_LAST_HILITE_VP*)
          (setq vpName (vl-string-trim " " vpName)))
        (setq vpPoly (halff:get-vp-poly-by-name vpName))
        (if (not vpPoly)
          (princ (strcat "\nX Viewport not found: " vpName))
          (progn
            (setq payRaw (getstring T
                           "\nEnter PAY ITEM to highlight (filtered to VP): "))
            (if (and payRaw (/= (vl-string-trim " " payRaw) ""))
              (progn
                (setq payKey (halff:resolve-paykey payRaw))
                (if (not payKey) (setq payKey (vl-string-trim " " payRaw)))
                (setq ss (halff:payitem->ss-vp payKey vpPoly))
                (if (and ss (> (sslength ss) 0))
                  (progn
                    (setq *HALFF_LAST_HILITE_VP* vpName)
                    (setq *HALFF_LAST_HILITE_PAY* payKey)
                    (sssetfirst nil ss)
                    (princ (strcat "\nOK Highlighted " (itoa (sslength ss))
                                   " entities for PAY ITEM "
                                   (halff:any->str payKey)
                                   " in VP " vpName)))
                  (princ (strcat "\nX Pay item not found in viewport "
                                 vpName ": " (halff:any->str payKey))))
                                 )
              (princ "\nX No pay item entered."))))
              )))
  (princ))

(defun c:QSEARCH (/ payRaw payKey rec ens idx n en obj prevObj ss cmd)
  (vl-load-com)
  (if (not *HALFF_PAYITEM_ENTS*)
    (princ "\nX No highlight data found. Run QRUN first.")
    (progn
      (setq payRaw (getstring T "\nEnter PAY ITEM to search: "))
      (if (or (not payRaw) (= (vl-string-trim " " payRaw) ""))
        (princ "\nX No pay item entered.")
        (progn
          (setq payKey (halff:resolve-paykey payRaw))
          (if (not payKey) (setq payKey (vl-string-trim " " payRaw)))
          (setq rec (assoc payKey *HALFF_PAYITEM_ENTS*))
          (if (and (not rec) (= (type payKey) 'STR) (wcmatch payKey "*#*"))
            (setq rec (assoc (rtos (atof payKey) 2 1) *HALFF_PAYITEM_ENTS*)))
          (if (not rec)
            (princ (strcat "\nX Pay item not found in this drawing: "
                           (halff:any->str payKey)))
            (progn
              (setq ens (cdr rec))
              (setq n (length ens))
              (if (or (not ens) (= n 0))
                (princ (strcat "\nX No cached entities for pay item: "
                               (halff:any->str payKey)))
                (progn
                  (setq idx 0)
                  (setq prevObj nil)
                  (while ens
                    (setq en (nth idx ens))
                    (if prevObj
                      (vl-catch-all-apply 'vla-Highlight
                                          (list prevObj :vlax-false)))
                    (if (and en (entget en))
                      (progn
                        (setq obj (vlax-ename->vla-object en))
                        (halff:zoom-to-ename en)
                        (vl-catch-all-apply 'vla-Highlight
                                            (list obj :vlax-true))
                        (setq prevObj obj)
                        (sssetfirst nil nil)
                        (setq ss (ssadd))
                        (ssadd en ss)
                        (sssetfirst nil ss)
                        (princ (strcat "\n[" (itoa (1+ idx)) "/" (itoa n)
                                       "] PAY ITEM "
                                       (halff:any->str payKey)
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
                             (setq ens nil))))
                             )
                      (setq idx
                        (if (< idx (1- n)) (1+ idx) 0))))
                        ))))
                        ))))
  (princ))

(defun c:QSEARCHVP (/ vpName vpPoly payRaw payKey rec ens idx n en obj prevObj ss cmd vpFilteredEnts)
  (if (not *HALFF_PAYITEM_ENTS*)
    (princ "\nNo entity cache. Run QRUN first.")
    (progn
      (setq vpName (getstring "\nEnter viewport name: "))
      (if (or (not vpName) (= vpName ""))
        (princ "\nCancelled.")
        (progn
          (setq vpPoly (halff:get-vp-poly-by-name vpName))
          (if (not vpPoly)
            (princ (strcat "\nViewport not found: " vpName))
            (progn
              (setq payRaw (getstring T
                             (strcat "\nEnter pay item to search in "
                                     vpName ": ")))
              (if (or (not payRaw) (= payRaw ""))
                (princ "\nCancelled.")
                (progn
                  (setq payKey (halff:resolve-paykey payRaw))
                  (setq rec (assoc payKey *HALFF_PAYITEM_ENTS*))
                  (if (not rec)
                    (princ (strcat "\nNo cached entities for pay item: "
                                   payRaw))
                    (progn
                      (setq ens (cdr rec))
                      (setq vpFilteredEnts '())
                      (foreach en ens
                        (if (and (entget en)
                                 (halff:vla-in-poly
                                   (vlax-ename->vla-object en) vpPoly))
                          (setq vpFilteredEnts
                            (append vpFilteredEnts (list en))))
                            )
                      (if (= (length vpFilteredEnts) 0)
                        (princ (strcat "\nNo entities found for " payRaw
                                       " in viewport " vpName))
                        (progn
                          (setq idx 0 n (length vpFilteredEnts))
                          (princ (strcat "\nFound " (itoa n) " entities in "
                                         vpName
                                         ". Use (N)ext, (P)revious, or (Q)uit."))
                          (setq prevObj nil)
                          (while (and (>= idx 0) (< idx n))
                            (setq en (nth idx vpFilteredEnts))
                            (if prevObj
                              (vl-catch-all-apply 'vla-Highlight
                                                  (list prevObj :vlax-false)))
                            (setq obj (vlax-ename->vla-object en))
                            (vla-Highlight obj :vlax-true)
                            (halff:zoom-to-ename en)
                            (setq ss (ssadd)) (ssadd en ss) (sssetfirst nil ss)
                            (princ (strcat "\n[" (itoa (1+ idx)) "/" (itoa n)
                                           "] " (halff:any->str payKey)
                                           " in " vpName
                                           " - (N)ext / (P)revious / (Q)uit: "))
                            (setq cmd (strcase (getstring)))
                            (cond
                              ((or (= cmd "N") (= cmd "")) (setq idx (1+ idx)))
                              ((= cmd "P") (setq idx (1- idx)))
                              ((= cmd "Q") (setq idx -1))
                              (T (princ "\nInvalid input.")))
                            (setq prevObj obj))
                          (if prevObj
                            (vl-catch-all-apply 'vla-Highlight
                                                (list prevObj :vlax-false)))
                          (princ "\nSearch ended."))))
                          ))))
                          ))))
                          )
  (princ))

;; -----------------------------------------------------------------------
;; QCIVIL3DPROBE -- iterate model space and report all Civil3D-like objects:
;;   DXF entity type, VLA ObjectName, and StyleName if available.
;;   Run this in the drawing that contains your pipes/structures to
;;   confirm the ObjectName and StyleName values QRUN needs to match.
(defun c:QCIVIL3DPROBE (/ ms obj oname en dxftype styleObj styleVal lenVal cnt)
  (vl-load-com)
  (setq ms (vla-get-ModelSpace
             (vla-get-ActiveDocument (vlax-get-acad-object))))
  (setq cnt 0)
  (princ "\nQCIVIL3DPROBE: scanning model space for Civil3D objects...")
  (vlax-for obj ms
    (setq en (vl-catch-all-apply 'vlax-vla-object->ename (list obj)))
    (setq dxftype
      (if (not (vl-catch-all-error-p en))
        (cdr (assoc 0 (entget en)))
        "?"))
    (setq oname (vl-catch-all-apply 'vla-get-ObjectName (list obj)))
    (if (vl-catch-all-error-p oname) (setq oname "?"))
    ;; Report anything with AECC in type/name, or non-standard AutoCAD entities
    (if (or (vl-string-search "AECC" (strcase dxftype))
            (vl-string-search "AECC" (strcase oname))
            (vl-string-search "PIPE" (strcase dxftype))
            (vl-string-search "STRUCT" (strcase dxftype)))
      (progn
        ;; Style via Style object -> Name
        (setq styleObj (vl-catch-all-apply 'vlax-get-property (list obj "Style")))
        (if (not (vl-catch-all-error-p styleObj))
          (setq styleVal (vl-catch-all-apply 'vlax-get-property (list styleObj "Name")))
          (setq styleVal nil))
        (if (or (null styleVal) (vl-catch-all-error-p styleVal))
          (setq styleVal "N/A"))
        ;; Length probe: Length2D -> Length3D -> Length
        (setq lenVal (vl-catch-all-apply 'vlax-get-property (list obj "Length2D")))
        (if (or (vl-catch-all-error-p lenVal) (null lenVal))
          (setq lenVal (vl-catch-all-apply 'vlax-get-property (list obj "Length3D"))))
        (if (or (vl-catch-all-error-p lenVal) (null lenVal))
          (setq lenVal (vl-catch-all-apply 'vlax-get-property (list obj "Length"))))
        (if (or (vl-catch-all-error-p lenVal) (null lenVal))
          (setq lenVal "N/A")
          (setq lenVal (rtos lenVal 2 4)))
        (princ (strcat "\n  DXF=" dxftype
                       "  ObjectName=" oname
                       "  Style.Name=" (vl-princ-to-string styleVal)
                       "  Length=" lenVal))
        (setq cnt (1+ cnt)))))
  (princ (strcat "\n\nFound " (itoa cnt) " Civil3D object(s)."))
  (princ "\nIf count=0, try QCIVIL3DPROBE2 to dump ALL entity types.")
  (princ))

;; List every unique Style.Name found on PIPE and STRUCTURE Civil3D objects,
;; with a count per style.  Run this in the target drawing to confirm that
;; the style names in your mapping file exactly match what Civil3D returns.
(defun c:QLISTPIPESTYLES
       (/ ms obj oname en dxftype styleObj styleVal descVal rows pair key)
  (vl-load-com)
  (setq ms (vla-get-ModelSpace
             (vla-get-ActiveDocument (vlax-get-acad-object))))
  (setq rows '())
  (princ "\nQLISTPIPESTYLES: collecting style+description from PIPE and STRUCTURE objects...")
  (vlax-for obj ms
    (setq en (vl-catch-all-apply 'vlax-vla-object->ename (list obj)))
    (setq dxftype
      (if (not (vl-catch-all-error-p en))
        (strcase (cdr (assoc 0 (entget en))))
        ""))
    (setq oname (vl-catch-all-apply 'vla-get-ObjectName (list obj)))
    (if (vl-catch-all-error-p oname) (setq oname ""))
    (if (or (vl-string-search "PIPE"   dxftype)
            (vl-string-search "STRUCT" dxftype)
            (vl-string-search "pipe"   (strcase oname T))
            (vl-string-search "struct" (strcase oname T)))
      (progn
        (setq styleObj (vl-catch-all-apply 'vlax-get-property (list obj "Style")))
        (setq styleVal
          (if (not (vl-catch-all-error-p styleObj))
            (vl-catch-all-apply 'vlax-get-property (list styleObj "Name"))
            nil))
        (if (or (null styleVal) (vl-catch-all-error-p styleVal))
          (setq styleVal "<no style>")
          (setq styleVal (vl-princ-to-string styleVal)))
        (setq descVal (vl-catch-all-apply 'vlax-get-property (list obj "Description")))
        (if (or (null descVal) (vl-catch-all-error-p descVal) (= descVal ""))
          (setq descVal "<no description>")
          (setq descVal (vl-princ-to-string descVal)))
        (setq key (strcat styleVal "|||" descVal))
        (setq pair (assoc key rows))
        (if pair
          (setq rows (subst (list key styleVal descVal (1+ (nth 3 pair))) pair rows))
          (setq rows (cons (list key styleVal descVal 1) rows))))))
  (princ (strcat "\n\nFound " (itoa (length rows)) " unique style+description combination(s):"))
  (foreach p (vl-sort rows '(lambda (a b) (< (nth 3 b) (nth 3 a))))
    (princ (strcat "\n  [" (itoa (nth 3 p)) "x]"
                   "  Style: " (nth 1 p)
                   "  |  Desc: " (nth 2 p))))
  (princ "\n\nUse Style and/or Description values exactly as shown in your mapping file.")
  (princ))

;; Fallback: dump every unique DXF entity type found in model space
(defun c:QCIVIL3DPROBE2 (/ ms obj en dxftype seen)
  (vl-load-com)
  (setq ms (vla-get-ModelSpace
             (vla-get-ActiveDocument (vlax-get-acad-object))))
  (setq seen '())
  (princ "\nQCIVIL3DPROBE2: all unique DXF entity types in model space:")
  (vlax-for obj ms
    (setq en (vl-catch-all-apply 'vlax-vla-object->ename (list obj)))
    (if (not (vl-catch-all-error-p en))
      (progn
        (setq dxftype (cdr (assoc 0 (entget en))))
        (if (not (member dxftype seen))
          (progn
            (setq seen (cons dxftype seen))
            (princ (strcat "\n  " dxftype)))))))
  (princ (strcat "\n\nTotal unique types: " (itoa (length seen))))
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
        (princ "\nOK No duplicates found."))))
  (princ))

;; ===========================================
;; FAILURE HIGHLIGHT COMMANDS
;; ===========================================

(defun c:QFAILHILITE (/ ss en)
  (if (not *HALFF_FAIL_ENTS*)
    (princ "\nX No failure data found. Run QRUN first.")
    (progn
      (setq ss (ssadd))
      (foreach en *HALFF_FAIL_ENTS*
        (if (and en (entget en))
          (ssadd en ss)))
      (if (> (sslength ss) 0)
        (progn
          (sssetfirst nil ss)
          (princ (strcat "\nOK Highlighted " (itoa (sslength ss))
                         " failure entities in current drawing.")))
        (princ "\nX No valid failure entities found in current drawing."))))
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
          (princ "\nQFAILSEARCH ended."))))
          )
  (princ))

;; ===========================================
;; ROUNDING COMMANDS
;; ===========================================

(defun c:QROUND (/)
  (cond
    ((= *HALFF_ROUND_MODE* "UP")
     (princ "\nCurrent rounding: ROUND UP (ceil) - Each quantity rounded up before adding"))
    ((= *HALFF_ROUND_MODE* "NEAREST")
     (princ "\nCurrent rounding: ROUND NEAREST - Each quantity rounded to nearest before adding"))
    (T (princ "\nCurrent rounding: UNKNOWN")))
  (princ))

(defun c:QROUNDUP (/)
  (setq *HALFF_ROUND_MODE* "UP")
  (princ "\nRounding set to: ROUND UP (ceil) - Each quantity rounded up before adding")
  (princ))

(defun c:QROUNDNEAR (/)
  (setq *HALFF_ROUND_MODE* "NEAREST")
  (princ "\nRounding set to: ROUND NEAREST - Each quantity rounded to nearest before adding")
  (princ))

(princ "\n+===============================================================+")
(princ "\n|  Halff QTO Labels v1.0                                        |")
(princ "\n|  • All geometry / layer / VP features from QTO v1.22          |")
(princ "\n|  • TEXT CONTENT column: count text entities by string match   |")
(princ "\n|  • Supports TEXT, MTEXT, MLEADER, Civil 3D label objects      |")
(princ "\n|  • Layer filter not required for text-count rows              |")
(princ "\n+===============================================================+")
(princ "\n")
(princ "\nViewport Commands:")
(princ "\n  QVPDEF          - Define viewports from polylines")
(princ "\n  QVPSET          - Set viewport definitions CSV path")
(princ "\n  QVPLOAD         - Load viewport definitions from CSV")
(princ "\n  QVPCROSSZERO    - Zero out crossing entities (default)")
(princ "\n  QVPCROSSCOUNT   - Count crossing entities")
(princ "\n  QVPEXCLUDE      - Set exclusion distance (default 1.0 ft)")
(princ "\n  QVPCROSS        - Show crossing detection settings")
(princ "\n")
(princ "\nRounding Commands:")
(princ "\n  QROUND          - Show current rounding mode")
(princ "\n  QROUNDUP        - Round up (ceil) each quantity - DEFAULT")
(princ "\n  QROUNDNEAR      - Round to nearest each quantity")
(princ "\n")
(princ "\nFull QTO Commands:")
(princ "\n  QMAPSET         - Set mapping Excel file path")
(princ "\n  QRUN            - Run quantity takeoff (geometry + text count)")
(princ "\n  QOPENREPORT     - Open output CSV")
(princ "\n")
(princ "\nHighlighting Commands:")
(princ "\n  QHILITE         - Highlight entities for a pay item (QRUN)")
(princ "\n  QHILITEVP       - Highlight entities in a viewport (QRUN)")
(princ "\n  QSEARCH         - Step through entities (QRUN)")
(princ "\n  QSEARCHVP       - Step through entities in viewport (QRUN)")
(princ "\n  QDUPLICATES     - Find duplicate objects")
(princ "\n  QFAILHILITE     - Highlight all failure entities (QRUN)")
(princ "\n  QFAILSEARCH     - Step through failure entities (QRUN)")
(princ "\n")