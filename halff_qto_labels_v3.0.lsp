;; =====================================================================
;; Halff QTO Labels v3.0
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
