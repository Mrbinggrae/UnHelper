; UnHelper per-user installer.
; APP_VERSION must be supplied by build_release.py:
;   makensis /DAPP_VERSION=0.1.0 UnHelper.nsi

Unicode True

!ifndef APP_VERSION
    !error "APP_VERSION is required. Run: python build_release.py un --installer"
!endif

!define APP_NAME     "UnHelper"
!define APP_EXE      "UnHelper.exe"
!define PUBLISHER    "Mrbinggrae"
!define DIST_DIR     "dist\UnHelper"
!define OUT_FILE     "release\UnHelper_Setup.exe"
!define APP_ICON     "assets\app-icon.ico"
!define REG_KEY      "Software\Microsoft\Windows\CurrentVersion\Uninstall\UnHelper"

Name "${APP_NAME} ${APP_VERSION}"
OutFile "${OUT_FILE}"
!if /FileExists "${APP_ICON}"
    Icon "${APP_ICON}"
    UninstallIcon "${APP_ICON}"
!endif
InstallDir "$LOCALAPPDATA\${APP_NAME}"
InstallDirRegKey HKCU "${REG_KEY}" "InstallLocation"
RequestExecutionLevel user
SetCompressor /SOLID lzma
ShowInstDetails show

Page directory
Page instfiles
UninstPage uninstConfirm
UninstPage instfiles

Section "Main" SEC_MAIN
    SetOutPath "$INSTDIR"
    File /r "${DIST_DIR}\*.*"

    CreateShortcut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\${APP_EXE}"
    CreateDirectory "$SMPROGRAMS\${APP_NAME}"
    CreateShortcut "$SMPROGRAMS\${APP_NAME}\${APP_NAME}.lnk" "$INSTDIR\${APP_EXE}" "" "$INSTDIR\${APP_EXE}"
    CreateShortcut "$SMPROGRAMS\${APP_NAME}\Uninstall.lnk" "$INSTDIR\Uninstall.exe"

    WriteUninstaller "$INSTDIR\Uninstall.exe"
    WriteRegStr   HKCU "${REG_KEY}" "DisplayName"      "${APP_NAME}"
    WriteRegStr   HKCU "${REG_KEY}" "DisplayVersion"   "${APP_VERSION}"
    WriteRegStr   HKCU "${REG_KEY}" "Publisher"        "${PUBLISHER}"
    WriteRegStr   HKCU "${REG_KEY}" "InstallLocation"  "$INSTDIR"
    WriteRegStr   HKCU "${REG_KEY}" "UninstallString"  '"$INSTDIR\Uninstall.exe"'
    WriteRegDWORD HKCU "${REG_KEY}" "NoModify"         1
    WriteRegDWORD HKCU "${REG_KEY}" "NoRepair"         1

    Exec '"$INSTDIR\${APP_EXE}"'
SectionEnd

Section "Uninstall"
    Delete "$DESKTOP\${APP_NAME}.lnk"
    RMDir /r "$SMPROGRAMS\${APP_NAME}"
    DeleteRegKey HKCU "${REG_KEY}"
    RMDir /r "$INSTDIR"
SectionEnd
