# winYAMB NSIS Configuration Script
# (d) by dUcA 2oo3.

# Customization ###################################################################
!define PRODUCT "winYAMB"
!define PRODUCT_LONG "Yamb Game for Windows"
!define VERSION "1.3"
!define YEAR "2oo3"
###################################################################################

# EXE-overHeads ###################################################################
VIAddVersionKey "ProductName" "${PRODUCT} Setup"
VIProductVersion "${VERSION}.0.0"
VIAddVersionKey "FileDescription" "${PRODUCT_LONG} Setup"
VIAddVersionKey "FileVersion" "${VERSION}"
VIAddVersionKey "LegalCopyright" "(d) by dUcA ${YEAR}."
###################################################################################

# Basic stuff #####################################################################
Name "${PRODUCT}"
Caption "${PRODUCT} ${VERSION} Setup"
OutFile "setup\${PRODUCT}.exe"
InstallDir "$PROGRAMFILES\${PRODUCT}"
InstallDirRegKey HKLM "Software\.dUcA\${PRODUCT}" ""
SetCompressor lzma
BrandingText "powered by Nullsoft Scriptable Install System"
SetFont "MS Sans Serif" "8"
CompletedText "Operations completed successfully!"
LicenseForceSelection checkbox "I &Accept"
#ShowInstDetails show
#ShowUninstDetails show
###################################################################################

# Install options #################################################################
InstType "Standard"
InstType "Full"
InstType /COMPONENTSONLYONCUSTOM
###################################################################################

# Include #########################################################################
!include "MUI.nsh"
!include "Sections.nsh"
###################################################################################

# Definitions #####################################################################
Var STARTMENU_FOLDER
!define MUI_FONT_HEADER "Arial"
!define MUI_FONTSIZE_HEADER "8"
!define MUI_FONTSTYLE_HEADER "700"
!define MUI_ICON "${PRODUCT}.ico"
!define MUI_UNICON "uninstall.ico"
!define MUI_CHECKBITMAP "modern.bmp"
!define MUI_COMPONENTSPAGE_SMALLDESC
###################################################################################

# Install pages ###################################################################
!insertmacro MUI_PAGE_LICENSE "license.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_STARTMENU Application $STARTMENU_FOLDER
!insertmacro MUI_PAGE_INSTFILES
  Page "custom" "FinishPage"
  ReserveFile "install.ini"
  Function FinishPage
	!insertmacro MUI_HEADER_TEXT "Installation Complete" "Setup was completed successfully."
	!insertmacro MUI_INSTALLOPTIONS_DISPLAY "install.ini"
  FunctionEnd
!insertmacro MUI_RESERVEFILE_INSTALLOPTIONS
###################################################################################

# Uninstall pages #################################################################
!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES
###################################################################################

# Language ########################################################################
!insertmacro MUI_LANGUAGE "English"
###################################################################################

# Install sections ################################################################
Section "!Program files (required)" SecProg
SectionIn 1 2 RO
  SetOutPath $INSTDIR
  File "winYAMB.exe"
  File "winYAMB.chm"
  WriteRegStr HKLM "Software\.dUcA\${PRODUCT}" "" "$INSTDIR"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "DisplayName" "winYAMB ${VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteUninstaller "uninstall.exe"
  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application
  	CreateDirectory "$SMPROGRAMS\$STARTMENU_FOLDER"
  	CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\winYAMB.lnk" "$INSTDIR\winYAMB.exe" '' "$INSTDIR\winYAMB.exe"
  	CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\winYAMB Help.lnk" "$INSTDIR\winYAMB.chm" '' "$INSTDIR\winYAMB.chm"
  	CreateShortCut "$SMPROGRAMS\$STARTMENU_FOLDER\Uninstall winYAMB.lnk" "$INSTDIR\uninstall.exe"
  !insertmacro MUI_STARTMENU_WRITE_END
SectionEnd

LangString DESC_SecProg ${LANG_ENGLISH} "${PRODUCT} program files (required executables: program, help system, uninstaller etc.)"

Section "Associate Game files" SecAssoc
SectionIn 1 2
  WriteRegStr HKCR ".wyb" "" "winYAMB.Game"
  WriteRegStr HKCR "winYAMB.Game" "" "winYAMB Game"
  WriteRegStr HKCR "winYAMB.Game\DefaultIcon" "" "$INSTDIR\winYAMB.exe,1"
  WriteRegStr HKCR "winYAMB.Game\shell" "" "&Open"
  WriteRegStr HKCR "winYAMB.Game\shell\&Open" "" ""
  WriteRegStr HKCR "winYAMB.Game\shell\&Open\command" "" '"$INSTDIR\winYAMB.exe" %1'
SectionEnd

LangString DESC_SecAssoc ${LANG_ENGLISH} "Associate winYAMB Game files with program executables in Windows environment"

Section "Sample Game files" SecSamp
SectionIn 2 3
  SetOutPath $INSTDIR
  File "Game1.wyb"
  File "Max.wyb"
SectionEnd

LangString DESC_SecSamp ${LANG_ENGLISH} "Sample winYAMB Game files (uncompleted game along with the maximum values for fields)"

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SecProg} $(DESC_SecProg)
  !insertmacro MUI_DESCRIPTION_TEXT ${SecAssoc} $(DESC_SecAssoc)
  !insertmacro MUI_DESCRIPTION_TEXT ${SecSamp} $(DESC_SecSamp)
!insertmacro MUI_FUNCTION_DESCRIPTION_END
###################################################################################

# Uninstall section ###############################################################
Section "Uninstall"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT}"
  DeleteRegKey HKLM "Software\.dUcA\${PRODUCT}"
  DeleteRegKey HKCR ".wyb"
  DeleteRegKey HKCR "winYAMB.Game"
  Delete $INSTDIR\*.*
  RMDir $INSTDIR
  !insertmacro MUI_STARTMENU_GETFOLDER Application $STARTMENU_FOLDER
  Delete "$SMPROGRAMS\$STARTMENU_FOLDER\*.*"
  RMDir "$SMPROGRAMS\$STARTMENU_FOLDER"
SectionEnd
###################################################################################

# Initialization function #########################################################
Function .onInit
  IfFileExists "$SYSDIR\msvbvm60.dll" 0 Abort
  !insertmacro MUI_INSTALLOPTIONS_EXTRACT "install.ini"
  Return
  Abort:
    MessageBox MB_OK|MB_ICONSTOP "Your system do not have Visual Basic 6 Runtime installed.$\n \
      $\n \
      Try to find installation file [vbrun60.exe] along with this$\n \
      distribution or from the place you got it in first place.$\n \
      $\n \
      - or -$\n \
      $\n \
      Try to download VB 6 Runtime from the Microsoft site:$\n \
      http://msdn.microsoft.com/downloads/$\n \
      $\n \
      This installer will now quit."
    Quit
FunctionEnd
###################################################################################

# Finish function #################################################################
Var INI_VALUE
Function .onGUIEnd
	!insertmacro MUI_INSTALLOPTIONS_READ $INI_VALUE "install.ini" "Field 2" "State"
	StrCmp $INI_VALUE 1 "" +2
   		Exec "$WINDIR\hh.exe $INSTDIR\${PRODUCT}.chm::/overview.html"
   	!insertmacro MUI_INSTALLOPTIONS_READ $INI_VALUE "install.ini" "Field 3" "State"
	StrCmp $INI_VALUE 1 "" +2
    	Exec "$INSTDIR\${PRODUCT}.exe"
FunctionEnd
###################################################################################