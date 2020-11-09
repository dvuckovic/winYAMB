VERSION 5.00
Begin VB.UserControl FieldBox 
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   Enabled         =   0   'False
   ScaleHeight     =   228.158
   ScaleMode       =   0  'User
   ScaleWidth      =   435
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000010&
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   375
   End
   Begin VB.Shape Box 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   257
      Left            =   0
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "FieldBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const p_def_Enabled = False
Const p_def_Shape = 0
Dim p_Enabled As Boolean
Dim p_Shape As Integer
Dim p_Text As String
Public Focus As Boolean
Event Click()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Private Sub Label_Click()
    If p_Enabled Then RaiseEvent Click
End Sub
Private Sub UserControl_GotFocus()
    Focus = True
    UserControl_Paint
End Sub
Private Sub UserControl_InitProperties()
    p_Enabled = p_def_Enabled
    p_Shape = p_def_Shape
    p_Text = vbNullString
End Sub
Private Sub DisEnab()
    If p_Enabled Then
        UserControl.Enabled = True
        Label.ForeColor = vbHighlight
        Box.BorderColor = vbHighlight
        Box.BackColor = &H80000014
    Else
        UserControl.Enabled = False
        Label.ForeColor = &H80000014
        Box.BorderColor = vbButtonShadow
        Box.BackColor = vbButtonFace
    End If
End Sub
Private Sub ChgShp()
    Box.Shape = p_Shape
End Sub
Private Sub ChgTxt()
    Label.Caption = p_Text
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If p_Enabled Then
        RaiseEvent KeyDown(KeyCode, Shift)
    End If
End Sub
Private Sub UserControl_LostFocus()
    Focus = False
    UserControl_Paint
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If p_Enabled Then RaiseEvent Click
End Sub
Public Sub UserControl_Paint()
    If p_Enabled Then
        If Focus Then
            Label.ForeColor = vbHighlight
            Box.BorderColor = &HC0
            Box.BackColor = &H80000014
        Else
            Label.ForeColor = vbHighlight
            Box.BorderColor = vbHighlight
            Box.BackColor = &H80000014
        End If
    Else
        Label.ForeColor = &H80000014
        Box.BorderColor = vbButtonShadow
        Box.BackColor = vbButtonFace
    End If
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    p_Enabled = PropBag.ReadProperty("Enabled", p_def_Enabled)
    p_Shape = PropBag.ReadProperty("Shape", p_def_Shape)
    p_Text = PropBag.ReadProperty("Text", vbNullString)
    DisEnab
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", p_Enabled, p_def_Enabled)
    Call PropBag.WriteProperty("Shape", p_Shape, p_def_Shape)
    Call PropBag.WriteProperty("Text", p_Text, vbNullString)
End Sub
Public Property Get Enabled() As Boolean
    Enabled = p_Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    p_Enabled = New_Enabled
    PropertyChanged "Enabled"
    DisEnab
End Property
Public Property Get Shape() As Integer
    Shape = p_Shape
End Property
Public Property Let Shape(ByVal New_Shape As Integer)
    p_Shape = New_Shape
    PropertyChanged "Shape"
    ChgShp
End Property
Public Property Get Text() As String
    Text = p_Text
End Property
Public Property Let Text(ByVal New_Text As String)
    p_Text = New_Text
    PropertyChanged "Text"
    ChgTxt
End Property
