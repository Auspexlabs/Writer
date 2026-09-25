; Included by Tauri's NSIS installer (bundle.windows.nsis.installerHooks in tauri.windows.conf.json). Tauri writes the
; file types itself (Software\Classes\.docx -> the ProgId named in bundle.fileAssociations, and so on), but Windows lets a
; user pick a default app only among Software\RegisteredApplications, each entry naming a Capabilities key that lists the
; app's ProgIds per extension. Without it Writer is missing from 设置 > 默认应用, and
; ms-settings:defaultapps?registeredAppUser=Writer (src/default_app.rs) has nothing to open. The ProgIds are read back
; from what Tauri just wrote, so the two never disagree. SHCTX is HKCU: the installer is per user.

!macro WRITER_FILE_TYPE EXT
  ReadRegStr $0 SHCTX "Software\Classes\.${EXT}" ""
  WriteRegStr SHCTX "Software\${MANUFACTURER}\${PRODUCTNAME}\Capabilities\FileAssociations" ".${EXT}" "$0"
!macroend

!macro NSIS_HOOK_POSTINSTALL
  Push $0
  WriteRegStr SHCTX "Software\${MANUFACTURER}\${PRODUCTNAME}\Capabilities" "ApplicationName" "${PRODUCTNAME}"
  WriteRegStr SHCTX "Software\${MANUFACTURER}\${PRODUCTNAME}\Capabilities" "ApplicationDescription" "${PRODUCTNAME}"
  !insertmacro WRITER_FILE_TYPE docx
  !insertmacro WRITER_FILE_TYPE xlsx
  !insertmacro WRITER_FILE_TYPE pptx
  !insertmacro WRITER_FILE_TYPE md
  !insertmacro WRITER_FILE_TYPE markdown
  !insertmacro WRITER_FILE_TYPE mm
  !insertmacro WRITER_FILE_TYPE pdf
  WriteRegStr SHCTX "Software\RegisteredApplications" "${PRODUCTNAME}" "Software\${MANUFACTURER}\${PRODUCTNAME}\Capabilities"
  Pop $0
!macroend

!macro NSIS_HOOK_POSTUNINSTALL
  DeleteRegValue SHCTX "Software\RegisteredApplications" "${PRODUCTNAME}"
  DeleteRegKey SHCTX "Software\${MANUFACTURER}\${PRODUCTNAME}\Capabilities"
!macroend
