VERSION 5.00
Begin VB.UserControl Hover 
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   ScaleHeight     =   615
   ScaleWidth      =   615
End
Attribute VB_Name = "Hover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZEL) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
 
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SIZEL
    cx As Long
    cy As Long
End Type

Enum EdgeType
    [None] = 0
    [Thick Raised] = &H5
    [Thick Sunken] = (&H2 Or &H8)
End Enum

Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Dim rRect As RECT
Dim fRect As RECT
Dim bHovering As Boolean
Dim bPressed As Boolean
Dim bFocused As Boolean
Dim nTextLeft As Long
Dim nTextTop As Long
Dim nControlHeight As Long
Dim nControlWidth As Long

Const m_def_BackColourNormal = vbButtonFace
Const m_def_BackColourHover = vbButtonFace
Const m_def_BackColourClick = vbButtonFace
Const m_def_ColourClick = vbHighlight
Const m_def_MouseHand = False
Const m_def_ColourHover = vbHighlight
Const m_def_ColourText = vbButtonText
Const m_def_BorderHover = &H5
Const m_def_BorderNormal = &H0
Const m_def_BorderClick = (&H2 Or &H8)
Const m_def_Enabled = True

Dim m_BorderHover As EdgeType
Dim m_BorderNormal As EdgeType
Dim m_BorderClick As EdgeType
Dim m_MouseHand As Boolean
Dim m_Enabled As Boolean
Dim m_Caption As String
Dim m_LinkToURL As Boolean
Dim m_URL As String
Dim m_Image As Picture

Event Click()
Private Sub UserControl_GotFocus()
    bFocused = True
    UserControl_Paint
End Sub
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_Caption = Ambient.DisplayName
    m_BorderHover = m_def_BorderHover
    m_BorderNormal = m_def_BorderNormal
    m_BorderClick = m_def_BorderClick
    Set m_Image = Nothing
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then bPressed = True: UserControl_Paint
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then bPressed = False: UserControl_Paint: RaiseEvent Click
End Sub
Private Sub UserControl_LostFocus()
    bFocused = False
    UserControl_Paint
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Temp As Boolean
    Call ReleaseCapture
    If (X < 0) Or (Y < 0) Or (X > nControlWidth) Or (Y > nControlHeight) Then
        Temp = False
    Else
        Temp = True
        Call SetCapture(UserControl.hwnd)
    End If
    If bHovering <> Temp Then
        If Button <> 0 Then bPressed = True
        bHovering = Temp
        If bHovering = False Then bPressed = False
        UserControl_Paint
    End If
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bPressed = True
    UserControl_Paint
End Sub
Private Sub UserControl_DblClick()
    bPressed = True
    UserControl_Paint
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bTemp As Boolean
    bTemp = bPressed
    bPressed = False
    bHovering = False
    UserControl_Paint
    If bTemp = True And m_Enabled = True Then
        DoEvents
        RaiseEvent Click
    End If
End Sub
Private Sub UserControl_Paint()
    Dim rct As RECT
    UserControl.Cls
    If m_Enabled = True Then
        UserControl.Enabled = True
        If m_Image Is Nothing Then
            If bFocused Then
                DrawFocusRect UserControl.hDC, fRect
            End If
            If bPressed Then
                DrawEdge UserControl.hDC, rRect, m_BorderClick, &H100F
                FigureTextSize
                UserControl.ForeColor = vbButtonText
                Call TextOut(UserControl.hDC, nTextLeft + 1, nTextTop + 1, m_Caption, Len(m_Caption))
            Else
                If bHovering Then
                    DrawEdge UserControl.hDC, rRect, m_BorderHover, &H100F
                Else
                    DrawEdge UserControl.hDC, rRect, m_BorderNormal, &H100F
                End If
                FigureTextSize
                UserControl.ForeColor = vbButtonText
                Call TextOut(UserControl.hDC, nTextLeft, nTextTop, m_Caption, Len(m_Caption))
            End If
        Else
            FigureTextSize
            UserControl.ForeColor = vbButtonText
            If bFocused Then
                DrawFocusRect UserControl.hDC, fRect
            End If
            If bPressed Then
                Call TextOut(UserControl.hDC, nTextLeft + 1, nTextTop + 1, m_Caption, Len(m_Caption))
                PaintPicture m_Image, 655, 55
                DrawEdge UserControl.hDC, rRect, m_BorderClick, &H100F
            Else
                Call TextOut(UserControl.hDC, nTextLeft, nTextTop, m_Caption, Len(m_Caption))
                PaintPicture m_Image, 640, 45
                If bHovering Then
                    DrawEdge UserControl.hDC, rRect, m_BorderHover, &H100F
                Else
                    DrawEdge UserControl.hDC, rRect, m_BorderNormal, &H100F
                End If
            End If
        End If
    Else
        UserControl.Enabled = False
        UserControl.ForeColor = vb3DHighlight
        bHovering = False
        bPressed = False
        bFocused = False
        FigureTextSize
        Call TextOut(UserControl.hDC, nTextLeft + 1, nTextTop + 1, m_Caption, Len(m_Caption))
        UserControl.ForeColor = vb3DShadow
        FigureTextSize
        Call TextOut(UserControl.hDC, nTextLeft, nTextTop, m_Caption, Len(m_Caption))
    End If
End Sub
Private Sub UserControl_Resize()
    rRect.Left = 0
    rRect.Top = 0
    rRect.Bottom = UserControl.Height \ Screen.TwipsPerPixelY
    rRect.Right = UserControl.Width \ Screen.TwipsPerPixelX
    fRect.Left = 3
    fRect.Top = 3
    fRect.Bottom = UserControl.Height \ Screen.TwipsPerPixelY - 3
    fRect.Right = UserControl.Width \ Screen.TwipsPerPixelX - 3
    nControlHeight = UserControl.Height
    nControlWidth = UserControl.Width
    FigureTextSize
    UserControl_Paint
End Sub
Private Sub FigureTextSize()
    Dim slTemp As SIZEL
    Call GetTextExtentPoint32(hDC, m_Caption, Len(m_Caption), slTemp)
    nTextLeft = (rRect.Right - slTemp.cx) \ 2
    nTextTop = (rRect.Bottom - slTemp.cy) \ 2
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Caption = PropBag.ReadProperty("Caption", UserControl.Ambient.DisplayName)
    m_BorderHover = PropBag.ReadProperty("BorderHover", m_def_BorderHover)
    m_BorderNormal = PropBag.ReadProperty("BorderNormal", m_def_BorderNormal)
    m_BorderClick = PropBag.ReadProperty("BorderClick", m_def_BorderClick)
    Set m_Image = PropBag.ReadProperty("Image", Nothing)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Caption", m_Caption, UserControl.Ambient.DisplayName)
    Call PropBag.WriteProperty("BorderHover", m_BorderHover, m_def_BorderHover)
    Call PropBag.WriteProperty("BorderNormal", m_BorderNormal, m_def_BorderNormal)
    Call PropBag.WriteProperty("BorderClick", m_BorderClick, m_def_BorderClick)
    Call PropBag.WriteProperty("Image", m_Image, Nothing)
End Sub
Public Property Get BorderHover() As EdgeType
    BorderHover = m_BorderHover
End Property
Public Property Let BorderHover(ByVal New_BorderHover As EdgeType)
    m_BorderHover = New_BorderHover
    PropertyChanged "BorderHover"
End Property
Public Property Get BorderNormal() As EdgeType
    BorderNormal = m_BorderNormal
End Property
Public Property Let BorderNormal(ByVal New_BorderNormal As EdgeType)
    m_BorderNormal = New_BorderNormal
    PropertyChanged "BorderNormal"
    UserControl_Paint
End Property
Public Property Get BorderClick() As EdgeType
    BorderClick = m_BorderClick
End Property
Public Property Let BorderClick(ByVal New_BorderClick As EdgeType)
    m_BorderClick = New_BorderClick
    PropertyChanged "BorderClick"
End Property
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    UserControl_Paint
End Property
Public Property Get Caption() As String
    Caption = m_Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    UserControl_Paint
End Property
Public Property Get Image() As Picture
    Set Image = m_Image
End Property
Public Property Set Image(ByVal New_Image As Picture)
    Set m_Image = New_Image
    PropertyChanged "Image"
    Refresh
End Property
