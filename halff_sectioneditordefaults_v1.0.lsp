;; =====================================================================
;; Halff SectionEditorDefaults (SED) v1.0
;;
;; Civil 3D's "View/Edit Corridor Section Options" dialog (command
;; AeccVecsOptions) is a plain modal PropertyGrid with no ActiveX/COM
;; exposure and no command-line/CMDDIA fallback -- there is no supported
;; API for setting its values from LISP. This routine drives the grid
;; with simulated keystrokes (WScript.Shell SendKeys) so the dialog
;; opens pre-filled with the team defaults captured below, instead of
;; you re-typing 20+ fields by hand every time before running the
;; Section Editor.
;;
;; HOW IT WORKS
;; - A small helper .vbs is written to %TEMP% and launched hidden and
;;   asynchronously (fire-and-forget) BEFORE the dialog opens.
;; - AutoLISP's (command) call blocks for the entire life of a modal
;;   dialog, so the keystrokes can't come from the LISP thread itself --
;;   they have to come from that independent helper process.
;; - The helper polls for a window titled "View/Edit Corridor Section
;;   Options" (WshShell.AppActivate) before typing anything. If that
;;   window never appears, it sends NOTHING -- it fails safe instead of
;;   typing into the wrong window (e.g. the drawing itself).
;; - SED never presses the dialog's OK button. It leaves every value
;;   filled in on screen for you to glance over and click OK yourself.
;;
;; ASSUMPTIONS BAKED INTO THE KEYSTROKES (verify on first run)
;; - The grid opens with the first category header ("View/Edit Options")
;;   already selected/highlighted.
;; - Up/Down arrow moves one row at a time through every visible row,
;;   including category headers (they are not skipped).
;; - Typing over a selected text/numeric row replaces its value outright
;;   (standard PropertyGrid behavior -- no F2 needed).
;; - Two-choice fields (Yes/No, On/Off) are opened with F4, then
;;   Home jumps to the first item in the list. This assumes the lists
;;   are ordered [No, Yes] and [Off, On]. If a Yes/No or On/Off field
;;   comes out backwards, that assumption is what to flip.
;; - Color swatch rows (grid line colors, axis color, text color,
;;   slider colors) are intentionally left untouched -- only skipped
;;   past with Down -- because a color picker sub-dialog can't be
;;   scripted blind. Set those once by hand; SED won't touch them.
;;
;; Run SED once before entering the Section Editor to push the drawing's
;; view/edit options back to the values below.
;; =====================================================================

(vl-load-com)

;; ===========================================
;; TARGET DEFAULT VALUES (edit here, not below)
;; ===========================================
(setq *HALFF_SED_VIEWSCALE*      "1.25")
(setq *HALFF_SED_REBUILD_ONEDIT* "Yes")   ;; Yes / No
(setq *HALFF_SED_FRONTCLIP*      "0.1")
(setq *HALFF_SED_BACKCLIP*       "-0.1")
(setq *HALFF_SED_APPLYVPCONFIG*  "On")    ;; On / Off
(setq *HALFF_SED_DISP_HGRID*     "Yes")
(setq *HALFF_SED_DISP_VGRID*     "Yes")
(setq *HALFF_SED_ADAPTIVEGRID*   "On")
(setq *HALFF_SED_MAJORLINEEVERY* "10")
(setq *HALFF_SED_HGRIDINTERVAL*  "1")
(setq *HALFF_SED_VGRIDINTERVAL*  "1")
(setq *HALFF_SED_DISP_CTRAXIS*   "Yes")
(setq *HALFF_SED_TEXTSTYLE*      "Standard")
(setq *HALFF_SED_TEXTHEIGHTPCT*  "2")
(setq *HALFF_SED_ANNOTATE_CTRAX* "No")
(setq *HALFF_SED_SLIDER_MULTIVP* "Yes")
(setq *HALFF_SED_CODESETSTYLE*   "LRT4_Assemblies_P")

;; ===========================================
;; KEYSTROKE HELPERS
;; ===========================================

;; One row down, type a value into a text/numeric field, commit.
(defun halff:sed-kv (val)
  (strcat "{DOWN}" val "{ENTER}"))

;; One row down, choose Yes/No from a dropdown opened with F4.
;; Assumes list order [No, Yes].
(defun halff:sed-kbool (target)
  (if (= (strcase target) "YES")
    "{DOWN}{F4}{HOME}{DOWN}{ENTER}"
    "{DOWN}{F4}{HOME}{ENTER}"))

;; One row down, choose On/Off from a dropdown opened with F4.
;; Assumes list order [Off, On].
(defun halff:sed-ktoggle (target)
  (if (= (strcase target) "ON")
    "{DOWN}{F4}{HOME}{DOWN}{ENTER}"
    "{DOWN}{F4}{HOME}{ENTER}"))

;; One row down, no edit (category header or color swatch left as-is).
(defun halff:sed-kskip ()
  "{DOWN}")

;; One row down, open a style-name dropdown with F4 and type the name.
(defun halff:sed-kstyle (val)
  (strcat "{DOWN}{F4}" val "{ENTER}"))

;; Full ordered walk of the property grid, top row to bottom row,
;; matching the "View/Edit Corridor Section Options" layout exactly.
(defun halff:sed-chunks ()
  (list
    (halff:sed-kv     *HALFF_SED_VIEWSCALE*)       ; Default View Scale
    (halff:sed-kbool  *HALFF_SED_REBUILD_ONEDIT*)  ; Rebuild on Edit
    (halff:sed-kv     *HALFF_SED_FRONTCLIP*)       ; Front Clip
    (halff:sed-kv     *HALFF_SED_BACKCLIP*)        ; Back Clip
    (halff:sed-ktoggle *HALFF_SED_APPLYVPCONFIG*)  ; Apply Viewport Configuration
    (halff:sed-kskip)                              ; [category] Grid Settings
    (halff:sed-kbool  *HALFF_SED_DISP_HGRID*)      ; Display Horizontal Grid
    (halff:sed-kbool  *HALFF_SED_DISP_VGRID*)      ; Display Vertical Grid
    (halff:sed-ktoggle *HALFF_SED_ADAPTIVEGRID*)   ; Adaptive Grid
    (halff:sed-kv     *HALFF_SED_MAJORLINEEVERY*)  ; Major Line Every
    (halff:sed-kv     *HALFF_SED_HGRIDINTERVAL*)   ; Horizontal Grid Interval
    (halff:sed-kv     *HALFF_SED_VGRIDINTERVAL*)   ; Vertical Grid Interval
    (halff:sed-kskip)                              ; Minor Grid Line Color (left as-is)
    (halff:sed-kskip)                              ; Major Grid Line Color (left as-is)
    (halff:sed-kbool  *HALFF_SED_DISP_CTRAXIS*)    ; Display Center Axis
    (halff:sed-kskip)                              ; Center Axis Color (left as-is)
    (halff:sed-kskip)                              ; [category] Grid Text Settings
    (halff:sed-kstyle *HALFF_SED_TEXTSTYLE*)       ; Text Style
    (halff:sed-kskip)                              ; Text Color (left as-is)
    (halff:sed-kv     *HALFF_SED_TEXTHEIGHTPCT*)   ; Text Height - Relative to Screen
    (halff:sed-kbool  *HALFF_SED_ANNOTATE_CTRAX*)  ; Annotate Center Axis
    (halff:sed-kskip)                              ; [category] Section Slider in Multiple...
    (halff:sed-kbool  *HALFF_SED_SLIDER_MULTIVP*)  ; Section Slider in Multiple Views
    (halff:sed-kskip)                              ; Horizontal Baseline Slider Color (left as-is)
    (halff:sed-kskip)                              ; Profile View Slider Color (left as-is)
    (halff:sed-kskip)                              ; [category] Default Styles
    (halff:sed-kstyle *HALFF_SED_CODESETSTYLE*)    ; Code Set Style
  ))

;; ===========================================
;; VBS HELPER GENERATION
;; ===========================================

(defun halff:sed-write-vbs (path chunks / f)
  (setq f (open path "w"))
  (write-line "Dim WshShell, found, i" f)
  (write-line "Set WshShell = CreateObject(\"WScript.Shell\")" f)
  (write-line "found = False" f)
  (write-line "For i = 1 To 30" f)
  (write-line "    If WshShell.AppActivate(\"View/Edit Corridor Section Options\") Then" f)
  (write-line "        found = True" f)
  (write-line "        Exit For" f)
  (write-line "    End If" f)
  (write-line "    WScript.Sleep 200" f)
  (write-line "Next" f)
  (write-line "" f)
  (write-line "If found Then" f)
  (write-line "    WScript.Sleep 300" f)
  (foreach chunk chunks
    (write-line (strcat "    WshShell.SendKeys \"" chunk "\"") f)
    (write-line "    WScript.Sleep 150" f))
  (write-line "End If" f)
  (close f)
  (princ))

;; ===========================================
;; COMMAND: SED / SectionEditorDefaults
;; ===========================================

(defun c:SED ( / tmp vbsPath shellObj err)
  (setq tmp (getenv "TEMP"))
  (if (or (not tmp) (= tmp "")) (setq tmp (getenv "TMP")))
  (if (or (not tmp) (= tmp "")) (setq tmp "C:\\Windows\\Temp"))
  (setq vbsPath (vl-filename-mktemp "halff_sed_" tmp ".vbs"))
  (halff:sed-write-vbs vbsPath (halff:sed-chunks))

  (setq shellObj (vl-catch-all-apply 'vlax-create-object (list "WScript.Shell")))
  (if (vl-catch-all-error-p shellObj)
    (progn
      (princ "\nSED: could not start the keystroke helper (WScript.Shell). Opening the dialog for manual entry.\n")
      (setq shellObj nil))
    (progn
      (setq err (vl-catch-all-apply 'vlax-invoke-method
                  (list shellObj 'Run
                        (strcat "wscript.exe //B //Nologo \"" vbsPath "\"")
                        0 :vlax-false)))
      (if (vl-catch-all-error-p err)
        (princ "\nSED: could not launch the keystroke helper. Opening the dialog for manual entry.\n"))))

  (princ "\nSED: opening View/Edit Corridor Section Options -- review the values, then click OK.\n")
  (command "_AeccVecsOptions")

  (if shellObj (vlax-release-object shellObj))
  (if (findfile vbsPath) (vl-file-delete vbsPath))
  (princ "\nSED: done -- confirm the values look right before running the Section Editor.\n")
  (princ))

(defun c:SectionEditorDefaults () (c:SED))

(princ "\nHalff SectionEditorDefaults loaded -- type SED to apply your Section Editor view/edit defaults.\n")
(princ)
