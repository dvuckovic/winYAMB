Attribute VB_Name = "CommonDialog"
Option Explicit
DefInt A-Z

Public Declare Function CommDlgExtendedError Lib "COMDLG32.DLL" () As Long
Public Declare Function GetOpenFileName Lib "COMDLG32.DLL" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetSaveFileName Lib "COMDLG32.DLL" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Type OPENFILENAME
 lStructSize        As Long
 hWndOwner          As Long
 hInstance          As Long
 lpStrFilter        As String
 lpStrCustomFilter  As String
 nMaxCustFilter     As Long
 nFilterIndex       As Long
 lpStrFile          As String
 nMaxFile           As Long
 lpStrFileTitle     As String
 nMaxFileTitle      As Long
 lpStrInitialDir    As String
 lpStrTitle         As String
 Flags              As Long
 nFileOffset        As Integer
 nFileExtension     As Integer
 lpStrDefExt        As String
 lCustData          As Long
 lpfnHook           As Long
 lpTemplateName     As String
End Type

Public Enum CDFileModes
 cdfmOpenFile
 cdfmOpenFileOrPrompt
 cdfmSaveFile
 cdfmSaveFileNoConfirm
End Enum

Public Enum CDFileFlags
 OFN_ALLOWMULTISELECT = &H200
 OFN_CREATEPROMPT = &H2000
 OFN_ENABLEHOOK = &H20
 OFN_ENABLETEMPLATE = &H40
 OFN_ENABLETEMPLATEHANDLE = &H80
 OFN_EXPLORER = &H80000
 OFN_EXTENSIONDIFFERENT = &H400
 OFN_FILEMUSTEXIST = &H1000
 OFN_HIDEREADONLY = &H4
 OFN_LONGNAMES = &H200000
 OFN_NOCHANGEDIR = &H8
 OFN_NODEREFERENCELINKS = &H100000
 OFN_NOLONGNAMES = &H40000
 OFN_NONETWORKBUTTON = &H20000
 OFN_NOREADONLYRETURN = &H8000
 OFN_NOTESTFILECREATE = &H10000
 OFN_NOVALIDATE = &H100
 OFN_OVERWRITEPROMPT = &H2
 OFN_PATHMUSTEXIST = &H800
 OFN_READONLY = &H1
 OFN_SHAREAWARE = &H4000
 OFN_SHAREFALLTHROUGH = 2
 OFN_SHAREWARN = 0
 OFN_SHARENOWARN = 1
 OFN_SHOWHELP = &H10
 OFS_MAXPATHNAME = 128
End Enum

Public Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_PATHMUSTEXIST Or OFN_FILEMUSTEXIST Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_PATHMUSTEXIST

Public OFN As OPENFILENAME
Public Function SelectFile(ByVal hWndOwner As Long, Optional ByVal Filter As String, Optional ByVal DefaultExtension As String, Optional ByVal FileMode As CDFileModes = cdfmOpenFile, Optional ByVal DialogCaption As String, Optional ByVal DefaultFilename As String, Optional ByVal DefaultPath As String, Optional FilterIDX As Long = 0, Optional MoreFlags As CDFileFlags) As String
 Dim R As Long, SP As Long, ShortSize As Long, Z As Long
 If InStr(DefaultFilename, "\") Then
  DefaultPath = GetPath(DefaultFilename)
  DefaultFilename = GetFile(DefaultFilename)
 End If
 If Len(DefaultPath) = 0 Then DefaultPath = CurDir$
 With OFN
  .lStructSize = Len(OFN)
  .hWndOwner = hWndOwner
  .hInstance = App.hInstance
  .lpStrFilter = Replace$(Filter, "|", Chr$(0)) & Chr$(0)
  .nFilterIndex = FilterIDX
  .lpStrFile = DefaultFilename & String$(257 - Len(DefaultFilename), 0)
  .nMaxFile = Len(.lpStrFile) - 1
  .lpStrFileTitle = .lpStrFile
  .nMaxFileTitle = .nMaxFile
  .lpStrDefExt = DefaultExtension & Chr$(0)
  .lpStrInitialDir = DefaultPath & Chr$(0)
  .lpStrTitle = DialogCaption & Chr$(0)
  If FileMode = cdfmSaveFile Or FileMode = cdfmSaveFileNoConfirm Then
   .Flags = OFS_FILE_SAVE_FLAGS
   If FileMode = cdfmSaveFileNoConfirm Then .Flags = .Flags Or OFN_OVERWRITEPROMPT
   R = GetSaveFileName(OFN)
  ElseIf FileMode = cdfmOpenFile Or FileMode = cdfmOpenFileOrPrompt Then
   .Flags = OFS_FILE_OPEN_FLAGS
   If FileMode = cdfmOpenFileOrPrompt Then .Flags = .Flags Or OFN_CREATEPROMPT
   R = GetOpenFileName(OFN)
  End If
  If R Then
   SP = InStr(.lpStrFile, Chr$(0))
   If SP Then .lpStrFile = Left$(.lpStrFile, SP - 1)
   SelectFile = Trim$(Replace$(.lpStrFile, Chr$(0), ""))
  Else
   Z = CommDlgExtendedError()
   If Z Then MsgBox "Unable to get filename(s)." & vbCr & vbCr & "CommDlgExtendedError returned " & Z, vbCritical
  End If
 End With
End Function
Private Function GetFile(ByVal PathAndFile As String) As String
 Dim R() As String
 If Len(PathAndFile) Then
  R() = Split(PathAndFile, "\")
  GetFile = R(UBound(R))
 End If
End Function
Private Function GetPath(ByVal Filename As String) As String
 Dim R() As String, P As String
 Dim I
 If InStr(Filename, "\") Then
  R() = Split(Filename, "\")
  For I = 0 To UBound(R) - 1
   P = P + R(I) + "\"
  Next
 End If
 GetPath = P
End Function
