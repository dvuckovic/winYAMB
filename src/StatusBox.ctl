VERSION 5.00
Begin VB.UserControl StatusBox 
   CanGetFocus     =   0   'False
   ClientHeight    =   255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ScaleHeight     =   255
   ScaleWidth      =   1515
   Begin VB.Line Liner 
      BorderColor     =   &H80000014&
      Index           =   3
      X1              =   0
      X2              =   1500
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Liner 
      BorderColor     =   &H80000014&
      Index           =   2
      X1              =   1500
      X2              =   1500
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line Liner 
      BorderColor     =   &H80000010&
      Index           =   4
      X1              =   0
      X2              =   0
      Y1              =   240
      Y2              =   0
   End
   Begin VB.Line Liner 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   0
      X2              =   1500
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label 
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   1425
   End
End
Attribute VB_Name = "StatusBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Const p_def_Align = vbLeftJustify
Const p_def_Status = vbNullString
Dim p_Status As String
Dim p_Align As AlignmentConstants
Private Sub UserControl_InitProperties()
    p_Status = p_def_Status
    p_Align = p_def_Align
End Sub
Private Sub UserControl_Paint()
    Label.Alignment = p_Align
    Label.Caption = p_Status
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    p_Status = PropBag.ReadProperty("Status", p_def_Status)
    p_Align = PropBag.ReadProperty("Align", p_def_Align)
End Sub
Private Sub UserControl_Resize()
    Liner(1).X2 = UserControl.Width
    Liner(2).X1 = UserControl.Width - 10
    Liner(2).X2 = UserControl.Width - 10
    Liner(3).X2 = UserControl.Width
    Label.Width = UserControl.Width - 80
    Label.Left = 45
    Label.Top = 30
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Status", p_Status, p_def_Status)
    Call PropBag.WriteProperty("Align", p_Align, p_def_Align)
End Sub
Public Property Get Status() As String
    Status = p_Status
End Property
Public Property Let Status(ByVal New_Status As String)
    p_Status = New_Status
    PropertyChanged "Status"
    UserControl_Paint
End Property
Public Property Get Align() As AlignmentConstants
    Align = p_Align
End Property
Public Property Let Align(ByVal New_Align As AlignmentConstants)
    p_Align = New_Align
    PropertyChanged "Align"
    UserControl_Paint
End Property
