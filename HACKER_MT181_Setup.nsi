Unicode true
!include "MUI2.nsh"

Name "HACKER M-T181 PC Search"
OutFile "HACKER_MT181_PC_CONNECT_CLEAN_WINDOWS_SETUP_V20.exe"
InstallDir "$LOCALAPPDATA\Programs\HackerMT181PC"
RequestExecutionLevel user
SetCompressor zlib
ShowInstDetails show
ShowUninstDetails show

!define MUI_ABORTWARNING
!define MUI_FINISHPAGE_RUN "$INSTDIR\HACKER_M_T181.exe"
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
!insertmacro MUI_LANGUAGE "English"

Section "Install"
  SetOutPath "$INSTDIR"
  SetOverwrite on
  File /r /x data /x __pycache__ "HACKER_MT181_WINDOWS_EXE\*"

  ; Empty data folder: no bundled user documents or cache.
  CreateDirectory "$INSTDIR\data"

  WriteUninstaller "$INSTDIR\Uninstall.exe"
  CreateDirectory "$SMPROGRAMS\HACKER M-T181 PC"
  CreateShortcut "$SMPROGRAMS\HACKER M-T181 PC\HACKER M-T181 PC.lnk" "$INSTDIR\HACKER_M_T181.exe"
  CreateShortcut "$DESKTOP\HACKER M-T181 PC.lnk" "$INSTDIR\HACKER_M_T181.exe"
  CreateShortcut "$SMPROGRAMS\HACKER M-T181 PC\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\HackerMT181PC" "DisplayName" "HACKER M-T181 PC Search"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\HackerMT181PC" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\HackerMT181PC" "InstallLocation" "$INSTDIR"
SectionEnd

Section "Uninstall"
  Delete "$DESKTOP\HACKER M-T181 PC.lnk"
  RMDir /r "$SMPROGRAMS\HACKER M-T181 PC"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\HackerMT181PC"
  RMDir /r "$INSTDIR\Lib"
  RMDir /r "$INSTDIR\tcl"
  RMDir /r "$INSTDIR\site-packages"
  RMDir /r "$INSTDIR\google_config"
  Delete "$INSTDIR\*.pyd"
  Delete "$INSTDIR\*.dll"
  Delete "$INSTDIR\*.exe"
  Delete "$INSTDIR\*.zip"
  Delete "$INSTDIR\*.py"
  Delete "$INSTDIR\*.bat"
  Delete "$INSTDIR\*.txt"
  Delete "$INSTDIR\*.cat"
  Delete "$INSTDIR\*.ico"
  Delete "$INSTDIR\python312._pth"
  Delete "$INSTDIR\Uninstall.exe"
  ; Keep data/, search cache and history so uninstall never destroys user data.
  RMDir "$INSTDIR"
SectionEnd
