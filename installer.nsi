; Installer for OfficeShrink
; James Singleton 2011 - unop.co.uk
;This work is licensed under a Creative Commons Attribution 3.0 Unported License.

;Set compression flag for compiler
SetCompressor /SOLID lzma
SetCompress auto

;Definitions
!define SHCNE_ASSOCCHANGED 0x8000000
!define SHCNF_IDLIST 0

; The name of the installer
Name "Office Shrink"

; The file to write
OutFile "OfficeShrinkInstaller.exe"

; The default installation directory
InstallDir $PROGRAMFILES\OfficeShrink

; Registry key to check for directory (so if you install again, it will 
; overwrite the old one automatically)
InstallDirRegKey HKLM "Software\OfficeShrink" "Install_Dir"

; Request application privileges for Windows Vista/7
RequestExecutionLevel admin

;--------------------------------

; Pages

Page license
Page components
Page directory
Page instfiles

UninstPage uninstConfirm
UninstPage instfiles

;--------------------------------

;Info screen, readme license etc.
LicenseData ReadMe.rtf
LicenseBkColor /windows
LicenseText "Welcome to the Office Shrink Installer. This program will install the Office Shrink tool on this computer. Please read the information below." "OK >"

; The stuff to install
Section "OfficeShrink"

  SectionIn RO
  
  ; Set output path to the installation directory.
  SetOutPath $INSTDIR
  
  ; Put batch files there
  File "OfficeShrink.bat"
  File "OfficeShrinkAndZip.bat"
  ; Documentation
  File "ReadMe.rtf"
  
  ; Add 7zip to path
  SetOutPath $WINDIR
  File "7za.exe"
  
  ; Write the installation path into the registry
  WriteRegStr HKLM SOFTWARE\OfficeShrink "Install_Dir" "$INSTDIR"
  
  ;Right click for office 2007 files
  ;$SENDTO as alternative and check type in tool?

  ;Word 2k7
  WriteRegStr HKCR "Word.Document.12\shell\shrink" "" "Shrink"
  WriteRegStr HKCR "Word.Document.12\shell\shrink\command" "" '"$INSTDIR\OfficeShrink.bat" "%1"'
  WriteRegStr HKCR "Word.Document.12\shell\shrink-zip" "" "Shrink and Zip"
  WriteRegStr HKCR "Word.Document.12\shell\shrink-zip\command" "" '"$INSTDIR\OfficeShrinkAndZip.bat"  "%1"'
  
  WriteRegStr HKCR "Word.DocumentMacroEnabled.12\shell\shrink" "" "Shrink"
  WriteRegStr HKCR "Word.DocumentMacroEnabled.12\shell\shrink\command" "" '"$INSTDIR\OfficeShrink.bat" "%1"'
  WriteRegStr HKCR "Word.DocumentMacroEnabled.12\shell\shrink-zip" "" "Shrink and Zip"
  WriteRegStr HKCR "Word.DocumentMacroEnabled.12\shell\shrink-zip\command" "" '"$INSTDIR\OfficeShrinkAndZip.bat"  "%1"'
  
  ;Excel 2k7 
  WriteRegStr HKCR "Excel.Sheet.12\shell\shrink" "" "Shrink"
  WriteRegStr HKCR "Excel.Sheet.12\shell\shrink\command" "" '"$INSTDIR\OfficeShrink.bat" "%1"'
  WriteRegStr HKCR "Excel.Sheet.12\shell\shrink-zip" "" "Shrink and Zip"
  WriteRegStr HKCR "Excel.Sheet.12\shell\shrink-zip\command" "" '"$INSTDIR\OfficeShrinkAndZip.bat"  "%1"'
  
  WriteRegStr HKCR "Excel.SheetMacroEnabled.12\shell\shrink" "" "Shrink"
  WriteRegStr HKCR "Excel.SheetMacroEnabled.12\shell\shrink\command" "" '"$INSTDIR\OfficeShrink.bat" "%1"'
  WriteRegStr HKCR "Excel.SheetMacroEnabled.12\shell\shrink-zip" "" "Shrink and Zip"
  WriteRegStr HKCR "Excel.SheetMacroEnabled.12\shell\shrink-zip\command" "" '"$INSTDIR\OfficeShrinkAndZip.bat"  "%1"'
    
  ;Powerpoint 2k7
  WriteRegStr HKCR "PowerPoint.Show.12\shell\shrink" "" "Shrink"
  WriteRegStr HKCR "PowerPoint.Show.12\shell\shrink\command" "" '"$INSTDIR\OfficeShrink.bat" "%1"'
  WriteRegStr HKCR "PowerPoint.Show.12\shell\shrink-zip" "" "Shrink and Zip"
  WriteRegStr HKCR "PowerPoint.Show.12\shell\shrink-zip\command" "" '"$INSTDIR\OfficeShrinkAndZip.bat"  "%1"'

  WriteRegStr HKCR "PowerPoint.ShowMacroEnabled.12\shell\shrink" "" "Shrink"
  WriteRegStr HKCR "PowerPoint.ShowMacroEnabled.12\shell\shrink\command" "" '"$INSTDIR\OfficeShrink.bat" "%1"'
  WriteRegStr HKCR "PowerPoint.ShowMacroEnabled.12\shell\shrink-zip" "" "Shrink and Zip"
  WriteRegStr HKCR "PowerPoint.ShowMacroEnabled.12\shell\shrink-zip\command" "" '"$INSTDIR\OfficeShrinkAndZip.bat"  "%1"'
  
  WriteRegStr HKCR "PowerPoint.SlideShow.12\shell\shrink" "" "Shrink"
  WriteRegStr HKCR "PowerPoint.SlideShow.12\shell\shrink\command" "" '"$INSTDIR\OfficeShrink.bat" "%1"'
  WriteRegStr HKCR "PowerPoint.SlideShow.12\shell\shrink-zip" "" "Shrink and Zip"
  WriteRegStr HKCR "PowerPoint.SlideShow.12\shell\shrink-zip\command" "" '"$INSTDIR\OfficeShrinkAndZip.bat"  "%1"'

  System::Call 'Shell32::SHChangeNotify(i ${SHCNE_ASSOCCHANGED}, i ${SHCNF_IDLIST}, i 0, i 0)'
  
  ; Write the uninstall keys for Windows
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OfficeShrink" "DisplayName" "OfficeShrink"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OfficeShrink" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OfficeShrink" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OfficeShrink" "NoRepair" 1
  WriteUninstaller "uninstall.exe"
  
SectionEnd

; Optional section (can be disabled by the user)
Section "Start Menu Shortcuts"

  CreateDirectory "$SMPROGRAMS\Office Shrink"
  CreateShortCut "$SMPROGRAMS\Office Shrink\Uninstall.lnk" "$INSTDIR\uninstall.exe" "" "$INSTDIR\uninstall.exe" 0
  ; Documentation
  CreateShortCut "$SMPROGRAMS\Office Shrink\About.lnk" "$INSTDIR\ReadMe.rtf" "" "$INSTDIR\ReadMe.rtf" 0
  
SectionEnd

;--------------------------------

; Uninstaller

Section "Uninstall"
  
  ; Remove registry keys
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\OfficeShrink"
  DeleteRegKey HKLM SOFTWARE\OfficeShrink

  ; Remove files and uninstaller
  Delete $INSTDIR\OfficeShrink.bat
  Delete $INSTDIR\OfficeShrinkAndZip.bat
  Delete $INSTDIR\uninstall.exe
  Delete $WINDIR\7za.exe
  ; Remove Documentation
  Delete $INSTDIR\ReadMe.rtf
  
  ; Remove right click options
  DeleteRegKey HKCR "Word.Document.12\shell\shrink" 
  DeleteRegKey HKCR "Word.Document.12\shell\shrink-zip" 
  
  DeleteRegKey HKCR "Word.DocumentMacroEnabled.12\shell\shrink" 
  DeleteRegKey HKCR "Word.DocumentMacroEnabled.12\shell\shrink-zip" 
  
  DeleteRegKey HKCR "Excel.Sheet.12\shell\shrink" 
  DeleteRegKey HKCR "Excel.Sheet.12\shell\shrink-zip" 
  
  DeleteRegKey HKCR "Excel.SheetMacroEnabled.12\shell\shrink" 
  DeleteRegKey HKCR "Excel.SheetMacroEnabled.12\shell\shrink-zip" 
  
  DeleteRegKey HKCR "PowerPoint.Show.12\shell\shrink" 
  DeleteRegKey HKCR "PowerPoint.Show.12\shell\shrink-zip" 
  
  DeleteRegKey HKCR "PowerPoint.ShowMacroEnabled.12\shell\shrink" 
  DeleteRegKey HKCR "PowerPoint.ShowMacroEnabled.12\shell\shrink-zip" 
  
  DeleteRegKey HKCR "PowerPoint.SlideShow.12\shell\shrink" 
  DeleteRegKey HKCR "PowerPoint.SlideShow.12\shell\shrink-zip" 

  System::Call 'Shell32::SHChangeNotify(i ${SHCNE_ASSOCCHANGED}, i ${SHCNF_IDLIST}, i 0, i 0)'

  ; Remove shortcuts, if any
  Delete "$SMPROGRAMS\Office Shrink\*.*"

  ; Remove directories used
  RMDir "$SMPROGRAMS\Office Shrink"
  RMDir "$INSTDIR"

SectionEnd