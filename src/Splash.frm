VERSION 5.00
Begin VB.Form Splash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   2955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame 
      Height          =   1020
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   2955
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   4
         Left            =   75
         TabIndex        =   7
         Top             =   795
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8.8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008FA0D1&
         Height          =   360
         Index           =   0
         Left            =   2325
         TabIndex        =   1
         Top             =   225
         Width           =   450
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "winYAMB v"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002144A3&
         Height          =   360
         Index           =   3
         Left            =   645
         TabIndex        =   2
         Top             =   225
         Width           =   1680
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8.8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   5
         Left            =   2340
         TabIndex        =   6
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(d) by dUcA 2oo3."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   2
         Left            =   1785
         TabIndex        =   4
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Link 
         AutoSize        =   -1  'True
         Caption         =   "http://duca.dnsalias.net"
         DragIcon        =   "Splash.frx":0000
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1455
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   795
         Width           =   1425
      End
      Begin VB.Image Iconist 
         Height          =   480
         Left            =   105
         Picture         =   "Splash.frx":0442
         Top             =   195
         Width           =   480
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "winYAMB v"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   360
         Index           =   1
         Left            =   660
         TabIndex        =   5
         Top             =   240
         Width           =   1680
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    Unload Me
End Sub
Private Sub Form_Deactivate()
    If Splash.Enabled Then Me.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub
Private Sub Form_Load()
    Label(0).Caption = App.Major & "." & App.Minor
    Label(5).Caption = App.Major & "." & App.Minor
    Me.Refresh
End Sub
Private Sub Frame_Click()
    Unload Me
End Sub
Private Sub Iconist_Click()
    Unload Me
End Sub
Private Sub Label_Click(Index As Integer)
    Unload Me
End Sub
Private Sub Link_DragDrop(Source As Control, X As Single, Y As Single)
    If Source Is Link Then
        With Link
            .ForeColor = vbBlack
            Call ShellExecute(0&, vbNullString, .Caption, vbNullString, vbNullString, vbNormalFocus)
        End With
    End If
End Sub
Private Sub Link_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If State = vbLeave Then
        With Link
            .Drag vbEndDrag
            .ForeColor = vbBlack
        End With
    End If
End Sub
Private Sub Link_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With Link
        .ForeColor = vbBlue
        .Drag vbBeginDrag
    End With
End Sub
