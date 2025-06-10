!define APPNAME "SlidesTranscriber"
!define EXE_NAME "transcribe-slides.exe"
!define OUTFILE "SlidesTranscriberSetup.exe"

OutFile "dist\${OUTFILE}"
InstallDir "$PROGRAMFILES\${APPNAME}"
RequestExecutionLevel user

Page directory
Page instfiles
Page uninstConfirm
Page uninstInstfiles

Section "Install"
  SetOutPath "$INSTDIR"
  File "dist\${EXE_NAME}"
  CreateShortcut "$DESKTOP\${APPNAME}.lnk" "$INSTDIR\${EXE_NAME}"
  CreateDirectory "$SMPROGRAMS\${APPNAME}"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "$INSTDIR\${EXE_NAME}"
  WriteUninstaller "$INSTDIR\uninstall.exe"
SectionEnd

Section "Uninstall"
  Delete "$DESKTOP\${APPNAME}.lnk"
  Delete "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk"
  RMDir "$SMPROGRAMS\${APPNAME}"
  Delete "$INSTDIR\${EXE_NAME}"
  Delete "$INSTDIR\uninstall.exe"
  RMDir "$INSTDIR"
SectionEnd
