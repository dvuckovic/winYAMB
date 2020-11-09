Attribute VB_Name = "Procedures"
Const VK_NUMLOCK = &H90
Const VK_SCROLL = &H91
Const VK_CAPITAL = &H14
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
Const VER_PLATFORM_WIN32_NT = 2
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const LOGPIXELSX = 88
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, ByVal lpType As String) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Sub NumLockHandle()
    Dim o As OSVERSIONINFO
    Dim NumLockState As Boolean
    o.dwOSVersionInfoSize = Len(o)
    GetVersionEx o
    Dim keys(0 To 255) As Byte
    GetKeyboardState keys(0)
    NumLockState = keys(VK_NUMLOCK)
    If NumLockState <> True Then
        If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            keys(VK_NUMLOCK) = 1
            SetKeyboardState keys(0)
        ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then
            keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
            keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        End If
    End If
End Sub
Public Function EncryptText(strText, strPwd)
    Dim I As Integer, c As Integer
    Dim strBuff As String
    strPwd = UCase$(strPwd)
    If Len(strPwd) Then
        For I = 1 To Len(strText)
            c = Asc(Mid$(strText, I, 1))
            c = c + Asc(Mid$(strPwd, (I Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next I
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function
Public Function DecryptText(strText, strPwd)
    Dim I As Integer, c As Integer
    Dim strBuff As String
    strPwd = UCase$(strPwd)
    If Len(strPwd) Then
        For I = 1 To Len(strText)
            c = Asc(Mid$(strText, I, 1))
            c = c - Asc(Mid$(strPwd, (I Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next I
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function
Sub Main()
    If ScreenFont Then
Loading:
        Splash.Show
        Splash.Label(4).Visible = True
        Splash.Enabled = False
        Splash.Refresh
        NumLockHandle
        PlayWav vbNullString
        Load Table
        If regQuery_A_Key(HKEY_LOCAL_MACHINE, "Software\.dUcA\winYAMB", "Appearance") = "@" Then Table.mnuOval_Click
        Table.Caption = "winYAMB"
        Table.Show
        Table.InitFields
        Table.RefreshStatus
        Table.StuffMenu
        Table.LoadBestScores
        If Command$ <> vbNullString Then
            On Error Resume Next
            Table.OpenIt Command$
        End If
        Unload Splash
    Else
        If MsgBox("You should set your font size to 96 dpi" & Chr(10) & _
                  "in Display Properties before continuing" & Chr(10) & _
                  "because winYAMB may not display properly." & Chr(10) & _
                  "Would you like to continue with loading?", vbYesNo + vbQuestion, "winYAMB") = vbYes Then GoTo Loading
    End If
End Sub
Public Sub PlayWav(rName As String)
    Dim hRsrc As Long
    Dim hGlobal As Long
    hRsrc = FindResource(App.hInstance, rName, "WAVE")
    hGlobal = LoadResource(App.hInstance, hRsrc)
    sndPlaySound hGlobal, 5
End Sub
Private Function ScreenFont() As Boolean
    Dim hWndDesk As Long
    Dim hDCDesk As Long
    Dim logPix As Long
    hWndDesk = GetDesktopWindow()
    hDCDesk = GetDC(hWndDesk)
    logPix = GetDeviceCaps(hDCDesk, LOGPIXELSX)
    Call ReleaseDC(hWndDesk, hDCDesk)
    ScreenFont = logPix = 96
End Function
