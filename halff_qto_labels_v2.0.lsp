;; =====================================================================
;; Halff QTO Labels v2.0
;;
;; Built on Halff QTO Labels v1.0.  Adds TEXT QTY column support so
;; that QRUN can extract and sum quantities embedded in TEXT, MTEXT,
;; and MLEADER entities (e.g. "PROP. 45 LF STEEL ENCASEMENT").
;;
;; NEW in v2.0 (vs v1.0):
;; - Excel mapping supports a new "TEXT QTY" column (after TEXT CONTENT).
;;   When both TEXT CONTENT and TEXT QTY are filled in, the row is
;;   processed in text-qty mode: the routine finds all matching text
;;   entities, extracts the number at the placeholder position, and
;;   SUMS those numbers as the quantity.
;; - TEXT CONTENT uses a placeholder character (typically "#") where
;;   the quantity number appears in the drawing label.
;;   Example:  TEXT CONTENT = "PROP. # LF STEEL ENCASEMENT"
;;             TEXT QTY     = "#"
;;   A label reading "PROP. 45 LF STEEL ENCASEMENT" yields qty 45.
;;   Multiple matching labels are summed.
;; - VP breakdown: each extracted quantity is attributed to the
;;   viewport containing that entity, same as geometry mode.
;; - Rows with TEXT CONTENT but no TEXT QTY continue to work in
;;   text-count mode (count = 1 per matching entity), unchanged.
;; - All geometry / layer / VP logic from v1.0/v1.22 is unchanged.
;; =====================================================================
