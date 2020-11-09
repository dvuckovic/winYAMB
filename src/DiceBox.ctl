VERSION 5.00
Begin VB.UserControl DiceBox 
   BackColor       =   &H80000014&
   ClientHeight    =   810
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   720
   ScaleHeight     =   724.737
   ScaleMode       =   0  'User
   ScaleWidth      =   720
   Begin VB.PictureBox Dice 
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   15
      ScaleHeight     =   780
      ScaleWidth      =   690
      TabIndex        =   0
      Top             =   15
      Width           =   690
   End
End
Attribute VB_Name = "DiceBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const p_def_Selected = False
Dim p_Die As Picture
Dim p_Selected As Boolean
Event Click()
Private Sub Dice_Click()
    RaiseEvent Click
End Sub
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub
Private Sub UserControl_InitProperties()
    Set p_Die = Nothing
    p_Selected = p_def_Selected
End Sub
Private Sub Selecting()
    If p_Selected Then
        UserControl.BackColor = vbHighlight
    Else
        UserControl.BackColor = vb3DHighlight
    End If
End Sub
Private Sub PaintDice()
    Set Dice.Picture = p_Die
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    p_Selected = PropBag.ReadProperty("Selected", p_def_Selected)
    Set p_Die = PropBag.ReadProperty("Die", Nothing)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Selected", p_Selected, p_def_Selected)
    Call PropBag.WriteProperty("Die", p_Die, Nothing)
End Sub
Public Property Get Selected() As Boolean
    Selected = p_Selected
End Property
Public Property Let Selected(ByVal New_Selected As Boolean)
    p_Selected = New_Selected
    PropertyChanged "Selected"
    Selecting
End Property
Public Property Get Die() As Picture
    Set Die = p_Die
End Property
Public Property Set Die(ByVal New_Die As Picture)
    Set p_Die = New_Die
    PropertyChanged "Die"
    PaintDice
End Property
