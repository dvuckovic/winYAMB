VERSION 5.00
Begin VB.Form Table 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "winYAMB"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   8070
   Icon            =   "Table.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   5640
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   2835
      Index           =   2
      Left            =   5415
      OLEDropMode     =   1  'Manual
      TabIndex        =   166
      Top             =   -90
      Width           =   2655
      Begin winYAMB.DiceBox Dice 
         Height          =   810
         Index           =   1
         Left            =   150
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   255
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1429
      End
      Begin winYAMB.Hover TurnBtn 
         Height          =   315
         Left            =   75
         TabIndex        =   105
         Top             =   2445
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   556
         Caption         =   "Next Turn"
      End
      Begin winYAMB.Hover RollBtn 
         Height          =   315
         Left            =   75
         TabIndex        =   104
         Top             =   2085
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   556
         Caption         =   "Roll dice!"
      End
      Begin winYAMB.DiceBox Dice 
         Height          =   810
         Index           =   6
         Left            =   1770
         TabIndex        =   191
         TabStop         =   0   'False
         Top             =   1140
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1429
      End
      Begin winYAMB.DiceBox Dice 
         Height          =   810
         Index           =   2
         Left            =   960
         TabIndex        =   192
         TabStop         =   0   'False
         Top             =   255
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1429
      End
      Begin winYAMB.DiceBox Dice 
         Height          =   810
         Index           =   3
         Left            =   1770
         TabIndex        =   193
         TabStop         =   0   'False
         Top             =   255
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1429
      End
      Begin winYAMB.DiceBox Dice 
         Height          =   810
         Index           =   4
         Left            =   150
         TabIndex        =   194
         TabStop         =   0   'False
         Top             =   1140
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1429
      End
      Begin winYAMB.DiceBox Dice 
         Height          =   810
         Index           =   5
         Left            =   960
         TabIndex        =   195
         TabStop         =   0   'False
         Top             =   1140
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1429
      End
      Begin VB.Label LabHit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hit F2 to start a new game..."
         Height          =   195
         Left            =   330
         OLEDropMode     =   1  'Manual
         TabIndex        =   189
         Top             =   1005
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Shape Shape 
         BackColor       =   &H80000014&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000D&
         Height          =   1875
         Index           =   0
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   165
         Width           =   2535
      End
   End
   Begin VB.Frame Frame 
      Height          =   5340
      Index           =   0
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   111
      Top             =   -90
      Width           =   5400
      Begin VB.PictureBox IcoLabel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   210
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   165
         TabStop         =   0   'False
         ToolTipText     =   "Downward Column"
         Top             =   150
         Width           =   240
      End
      Begin VB.PictureBox SumLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Height          =   135
         Index           =   3
         Left            =   5070
         OLEDropMode     =   1  'Manual
         Picture         =   "Table.frx":0E42
         ScaleHeight     =   135
         ScaleWidth      =   90
         TabIndex        =   159
         TabStop         =   0   'False
         Top             =   195
         Width           =   90
      End
      Begin VB.PictureBox ColLabel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   105
         Index           =   3
         Left            =   3412
         OLEDropMode     =   1  'Manual
         Picture         =   "Table.frx":0E83
         ScaleHeight     =   105
         ScaleWidth      =   270
         TabIndex        =   155
         TabStop         =   0   'False
         ToolTipText     =   "Middle Column"
         Top             =   210
         Width           =   270
      End
      Begin VB.PictureBox ColLabel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   4
         Left            =   3990
         OLEDropMode     =   1  'Manual
         Picture         =   "Table.frx":0ECF
         ScaleHeight     =   195
         ScaleWidth      =   165
         TabIndex        =   154
         TabStop         =   0   'False
         ToolTipText     =   "TopBottom Column"
         Top             =   165
         Width           =   165
      End
      Begin VB.PictureBox ColLabel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   90
         Index           =   2
         Left            =   1890
         OLEDropMode     =   1  'Manual
         Picture         =   "Table.frx":0F1C
         ScaleHeight     =   90
         ScaleWidth      =   165
         TabIndex        =   136
         TabStop         =   0   'False
         ToolTipText     =   "Upward Column"
         Top             =   210
         Width           =   165
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   1
         Left            =   705
         TabIndex        =   0
         Top             =   450
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin VB.PictureBox ColLabel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   1
         Left            =   1365
         OLEDropMode     =   1  'Manual
         Picture         =   "Table.frx":0F5E
         ScaleHeight     =   195
         ScaleWidth      =   165
         TabIndex        =   129
         TabStop         =   0   'False
         ToolTipText     =   "Random Column"
         Top             =   165
         Width           =   165
      End
      Begin VB.PictureBox ColLabel 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   90
         Index           =   0
         Left            =   840
         OLEDropMode     =   1  'Manual
         Picture         =   "Table.frx":0FAB
         ScaleHeight     =   90
         ScaleWidth      =   165
         TabIndex        =   128
         TabStop         =   0   'False
         ToolTipText     =   "Downward Column"
         Top             =   210
         Width           =   165
      End
      Begin VB.PictureBox SumLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Height          =   135
         Index           =   2
         Left            =   270
         OLEDropMode     =   1  'Manual
         Picture         =   "Table.frx":0FED
         ScaleHeight     =   135
         ScaleWidth      =   90
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   5085
         Width           =   90
      End
      Begin VB.PictureBox SumLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Height          =   135
         Index           =   1
         Left            =   285
         OLEDropMode     =   1  'Manual
         Picture         =   "Table.frx":102E
         ScaleHeight     =   135
         ScaleWidth      =   90
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   3255
         Width           =   90
      End
      Begin VB.PictureBox SumLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Height          =   135
         Index           =   0
         Left            =   285
         OLEDropMode     =   1  'Manual
         Picture         =   "Table.frx":106F
         ScaleHeight     =   135
         ScaleWidth      =   90
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   2280
         Width           =   90
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   13
         Left            =   705
         TabIndex        =   12
         Top             =   4680
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   2
         Left            =   705
         TabIndex        =   1
         Top             =   735
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   3
         Left            =   705
         TabIndex        =   2
         Top             =   1020
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   4
         Left            =   705
         TabIndex        =   3
         Top             =   1305
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   5
         Left            =   705
         TabIndex        =   4
         Top             =   1590
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   6
         Left            =   705
         TabIndex        =   5
         Top             =   1875
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   7
         Left            =   705
         TabIndex        =   6
         Top             =   2565
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   8
         Left            =   705
         TabIndex        =   7
         Top             =   2850
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   9
         Left            =   705
         TabIndex        =   8
         Top             =   3540
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   10
         Left            =   705
         TabIndex        =   9
         Top             =   3825
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   11
         Left            =   705
         TabIndex        =   10
         Top             =   4110
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Down 
         Height          =   255
         Index           =   12
         Left            =   705
         TabIndex        =   11
         Top             =   4395
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   21
         Left            =   1230
         TabIndex        =   13
         Top             =   450
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   33
         Left            =   1230
         TabIndex        =   25
         Top             =   4680
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   22
         Left            =   1230
         TabIndex        =   14
         Top             =   735
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   23
         Left            =   1230
         TabIndex        =   15
         Top             =   1020
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   24
         Left            =   1230
         TabIndex        =   16
         Top             =   1305
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   25
         Left            =   1230
         TabIndex        =   17
         Top             =   1590
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   26
         Left            =   1230
         TabIndex        =   18
         Top             =   1875
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   27
         Left            =   1230
         TabIndex        =   19
         Top             =   2565
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   28
         Left            =   1230
         TabIndex        =   20
         Top             =   2850
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   29
         Left            =   1230
         TabIndex        =   21
         Top             =   3540
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   30
         Left            =   1230
         TabIndex        =   22
         Top             =   3825
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   31
         Left            =   1230
         TabIndex        =   23
         Top             =   4110
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Random 
         Height          =   255
         Index           =   32
         Left            =   1230
         TabIndex        =   24
         Top             =   4395
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   41
         Left            =   1755
         TabIndex        =   26
         Top             =   450
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   53
         Left            =   1755
         TabIndex        =   38
         Top             =   4680
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   42
         Left            =   1755
         TabIndex        =   27
         Top             =   735
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   43
         Left            =   1755
         TabIndex        =   28
         Top             =   1020
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   44
         Left            =   1755
         TabIndex        =   29
         Top             =   1305
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   45
         Left            =   1755
         TabIndex        =   30
         Top             =   1590
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   46
         Left            =   1755
         TabIndex        =   31
         Top             =   1875
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   47
         Left            =   1755
         TabIndex        =   32
         Top             =   2565
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   48
         Left            =   1755
         TabIndex        =   33
         Top             =   2850
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   49
         Left            =   1755
         TabIndex        =   34
         Top             =   3540
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   50
         Left            =   1755
         TabIndex        =   35
         Top             =   3825
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   51
         Left            =   1755
         TabIndex        =   36
         Top             =   4110
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Up 
         Height          =   255
         Index           =   52
         Left            =   1755
         TabIndex        =   37
         Top             =   4395
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   61
         Left            =   2280
         TabIndex        =   39
         Top             =   450
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   73
         Left            =   2280
         TabIndex        =   51
         Top             =   4680
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   62
         Left            =   2280
         TabIndex        =   40
         Top             =   735
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   63
         Left            =   2280
         TabIndex        =   41
         Top             =   1020
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   64
         Left            =   2280
         TabIndex        =   42
         Top             =   1305
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   65
         Left            =   2280
         TabIndex        =   43
         Top             =   1590
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   66
         Left            =   2280
         TabIndex        =   44
         Top             =   1875
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   67
         Left            =   2280
         TabIndex        =   45
         Top             =   2565
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   68
         Left            =   2280
         TabIndex        =   46
         Top             =   2850
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   69
         Left            =   2280
         TabIndex        =   47
         Top             =   3540
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   70
         Left            =   2280
         TabIndex        =   48
         Top             =   3825
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   71
         Left            =   2280
         TabIndex        =   49
         Top             =   4110
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Announce 
         Height          =   255
         Index           =   72
         Left            =   2280
         TabIndex        =   50
         Top             =   4395
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   81
         Left            =   2798
         TabIndex        =   52
         Top             =   450
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   93
         Left            =   2805
         TabIndex        =   64
         Top             =   4680
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   82
         Left            =   2805
         TabIndex        =   53
         Top             =   735
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   83
         Left            =   2805
         TabIndex        =   54
         Top             =   1020
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   84
         Left            =   2805
         TabIndex        =   55
         Top             =   1305
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   85
         Left            =   2805
         TabIndex        =   56
         Top             =   1590
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   86
         Left            =   2805
         TabIndex        =   57
         Top             =   1875
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   87
         Left            =   2805
         TabIndex        =   58
         Top             =   2565
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   88
         Left            =   2805
         TabIndex        =   59
         Top             =   2850
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   89
         Left            =   2805
         TabIndex        =   60
         Top             =   3540
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   90
         Left            =   2805
         TabIndex        =   61
         Top             =   3825
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   91
         Left            =   2805
         TabIndex        =   62
         Top             =   4110
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Hand 
         Height          =   255
         Index           =   92
         Left            =   2805
         TabIndex        =   63
         Top             =   4395
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   113
         Left            =   3330
         TabIndex        =   77
         Top             =   4680
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   102
         Left            =   3330
         TabIndex        =   66
         Top             =   735
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   103
         Left            =   3330
         TabIndex        =   67
         Top             =   1020
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   104
         Left            =   3330
         TabIndex        =   68
         Top             =   1305
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   105
         Left            =   3330
         TabIndex        =   69
         Top             =   1590
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   106
         Left            =   3330
         TabIndex        =   70
         Top             =   1875
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   107
         Left            =   3330
         TabIndex        =   71
         Top             =   2565
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   108
         Left            =   3330
         TabIndex        =   72
         Top             =   2850
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   109
         Left            =   3330
         TabIndex        =   73
         Top             =   3540
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   110
         Left            =   3330
         TabIndex        =   74
         Top             =   3825
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   111
         Left            =   3330
         TabIndex        =   75
         Top             =   4110
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   112
         Left            =   3330
         TabIndex        =   76
         Top             =   4395
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   121
         Left            =   3855
         TabIndex        =   78
         Top             =   450
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   133
         Left            =   3855
         TabIndex        =   90
         Top             =   4680
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   122
         Left            =   3855
         TabIndex        =   79
         Top             =   735
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   123
         Left            =   3855
         TabIndex        =   80
         Top             =   1020
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   124
         Left            =   3855
         TabIndex        =   81
         Top             =   1305
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   125
         Left            =   3855
         TabIndex        =   82
         Top             =   1590
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   126
         Left            =   3855
         TabIndex        =   83
         Top             =   1875
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   127
         Left            =   3855
         TabIndex        =   84
         Top             =   2565
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   128
         Left            =   3855
         TabIndex        =   85
         Top             =   2850
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   129
         Left            =   3855
         TabIndex        =   86
         Top             =   3540
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   130
         Left            =   3855
         TabIndex        =   87
         Top             =   3825
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   131
         Left            =   3855
         TabIndex        =   88
         Top             =   4110
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox UpDown 
         Height          =   255
         Index           =   132
         Left            =   3855
         TabIndex        =   89
         Top             =   4395
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Middle 
         Height          =   255
         Index           =   101
         Left            =   3330
         TabIndex        =   65
         Top             =   450
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   141
         Left            =   4380
         TabIndex        =   91
         Top             =   450
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   153
         Left            =   4380
         TabIndex        =   103
         Top             =   4680
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   142
         Left            =   4380
         TabIndex        =   92
         Top             =   735
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   143
         Left            =   4380
         TabIndex        =   93
         Top             =   1020
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   144
         Left            =   4380
         TabIndex        =   94
         Top             =   1305
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   145
         Left            =   4380
         TabIndex        =   95
         Top             =   1590
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   146
         Left            =   4380
         TabIndex        =   96
         Top             =   1875
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   147
         Left            =   4380
         TabIndex        =   97
         Top             =   2565
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   148
         Left            =   4380
         TabIndex        =   98
         Top             =   2850
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   149
         Left            =   4380
         TabIndex        =   99
         Top             =   3540
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   150
         Left            =   4380
         TabIndex        =   100
         Top             =   3825
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   151
         Left            =   4380
         TabIndex        =   101
         Top             =   4110
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin winYAMB.FieldBox Max 
         Height          =   255
         Index           =   152
         Left            =   4380
         TabIndex        =   102
         Top             =   4395
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   450
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   163
         Left            =   4890
         OLEDropMode     =   1  'Manual
         TabIndex        =   163
         Top             =   5055
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   162
         Left            =   4890
         OLEDropMode     =   1  'Manual
         TabIndex        =   162
         Top             =   3225
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   161
         Left            =   4890
         OLEDropMode     =   1  'Manual
         TabIndex        =   161
         Top             =   2250
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   141
         Left            =   4365
         OLEDropMode     =   1  'Manual
         TabIndex        =   160
         Top             =   2250
         Width           =   435
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   64
         X1              =   4860
         X2              =   4860
         Y1              =   3465
         Y2              =   4995
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   13
         X1              =   15
         X2              =   4875
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   12
         X1              =   30
         X2              =   4845
         Y1              =   3495
         Y2              =   3495
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   19
         X1              =   645
         X2              =   645
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   23
         X1              =   1170
         X2              =   1170
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   27
         X1              =   1695
         X2              =   1695
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   32
         X1              =   2220
         X2              =   2220
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   39
         X1              =   2745
         X2              =   2745
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   44
         X1              =   3270
         X2              =   3270
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   51
         X1              =   3795
         X2              =   3795
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   56
         X1              =   4320
         X2              =   4320
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   8
         X1              =   15
         X2              =   4845
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   30
         X2              =   4870
         Y1              =   405
         Y2              =   405
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   65
         X1              =   4845
         X2              =   4845
         Y1              =   3495
         Y2              =   4980
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   63
         X1              =   4845
         X2              =   4845
         Y1              =   2505
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   62
         X1              =   4860
         X2              =   4860
         Y1              =   2490
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   61
         X1              =   4845
         X2              =   4845
         Y1              =   105
         Y2              =   2160
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   60
         X1              =   4860
         X2              =   4860
         Y1              =   105
         Y2              =   2175
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "M"
         Height          =   195
         Index           =   15
         Left            =   4530
         OLEDropMode     =   1  'Manual
         TabIndex        =   158
         ToolTipText     =   "Maximum Column"
         Top             =   165
         Width           =   135
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   142
         Left            =   4365
         OLEDropMode     =   1  'Manual
         TabIndex        =   157
         Top             =   3225
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   143
         Left            =   4365
         OLEDropMode     =   1  'Manual
         TabIndex        =   156
         Top             =   5055
         Width           =   435
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   59
         X1              =   4335
         X2              =   4335
         Y1              =   105
         Y2              =   2175
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   58
         X1              =   4320
         X2              =   4320
         Y1              =   105
         Y2              =   2160
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   57
         X1              =   4335
         X2              =   4335
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   55
         X1              =   4335
         X2              =   4335
         Y1              =   3495
         Y2              =   4995
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   54
         X1              =   4320
         X2              =   4320
         Y1              =   3495
         Y2              =   4980
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   123
         Left            =   3840
         OLEDropMode     =   1  'Manual
         TabIndex        =   153
         Top             =   5055
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   121
         Left            =   3840
         OLEDropMode     =   1  'Manual
         TabIndex        =   152
         Top             =   2250
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   122
         Left            =   3840
         OLEDropMode     =   1  'Manual
         TabIndex        =   151
         Top             =   3225
         Width           =   435
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   53
         X1              =   3795
         X2              =   3795
         Y1              =   3495
         Y2              =   4980
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   52
         X1              =   3810
         X2              =   3810
         Y1              =   3495
         Y2              =   4995
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   50
         X1              =   3810
         X2              =   3810
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   49
         X1              =   3795
         X2              =   3795
         Y1              =   105
         Y2              =   2160
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   48
         X1              =   3810
         X2              =   3810
         Y1              =   105
         Y2              =   2175
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   102
         Left            =   3315
         OLEDropMode     =   1  'Manual
         TabIndex        =   150
         Top             =   3225
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   101
         Left            =   3315
         OLEDropMode     =   1  'Manual
         TabIndex        =   149
         Top             =   2250
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   103
         Left            =   3315
         OLEDropMode     =   1  'Manual
         TabIndex        =   148
         Top             =   5055
         Width           =   435
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   47
         X1              =   3285
         X2              =   3285
         Y1              =   105
         Y2              =   2175
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   46
         X1              =   3270
         X2              =   3270
         Y1              =   105
         Y2              =   2160
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   45
         X1              =   3285
         X2              =   3285
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   43
         X1              =   3285
         X2              =   3285
         Y1              =   3495
         Y2              =   4995
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   42
         X1              =   3270
         X2              =   3270
         Y1              =   3495
         Y2              =   4980
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   83
         Left            =   2790
         OLEDropMode     =   1  'Manual
         TabIndex        =   147
         Top             =   5055
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   81
         Left            =   2790
         OLEDropMode     =   1  'Manual
         TabIndex        =   146
         Top             =   2250
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   82
         Left            =   2790
         OLEDropMode     =   1  'Manual
         TabIndex        =   145
         Top             =   3225
         Width           =   435
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "H"
         Height          =   195
         Index           =   14
         Left            =   2955
         OLEDropMode     =   1  'Manual
         TabIndex        =   144
         ToolTipText     =   "Hand Column"
         Top             =   165
         Width           =   120
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   62
         Left            =   2280
         OLEDropMode     =   1  'Manual
         TabIndex        =   143
         Top             =   3225
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   61
         Left            =   2280
         OLEDropMode     =   1  'Manual
         TabIndex        =   142
         Top             =   2250
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   63
         Left            =   2280
         OLEDropMode     =   1  'Manual
         TabIndex        =   141
         Top             =   5055
         Width           =   435
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   41
         X1              =   2745
         X2              =   2745
         Y1              =   3495
         Y2              =   4980
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   40
         X1              =   2760
         X2              =   2760
         Y1              =   3495
         Y2              =   4995
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   38
         X1              =   2760
         X2              =   2760
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   37
         X1              =   2745
         X2              =   2745
         Y1              =   105
         Y2              =   2160
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   36
         X1              =   2760
         X2              =   2760
         Y1              =   105
         Y2              =   2175
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "A"
         Height          =   195
         Index           =   13
         Left            =   2445
         OLEDropMode     =   1  'Manual
         TabIndex        =   140
         ToolTipText     =   "Announce Column"
         Top             =   165
         Width           =   105
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   43
         Left            =   1755
         OLEDropMode     =   1  'Manual
         TabIndex        =   139
         Top             =   5055
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   41
         Left            =   1755
         OLEDropMode     =   1  'Manual
         TabIndex        =   138
         Top             =   2250
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   42
         Left            =   1755
         OLEDropMode     =   1  'Manual
         TabIndex        =   137
         Top             =   3225
         Width           =   435
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   35
         X1              =   2235
         X2              =   2235
         Y1              =   105
         Y2              =   2175
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   34
         X1              =   2220
         X2              =   2220
         Y1              =   105
         Y2              =   2160
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   33
         X1              =   2235
         X2              =   2235
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   31
         X1              =   2235
         X2              =   2235
         Y1              =   3495
         Y2              =   4995
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   30
         X1              =   2220
         X2              =   2220
         Y1              =   3495
         Y2              =   4980
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   29
         X1              =   1695
         X2              =   1695
         Y1              =   3495
         Y2              =   4980
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   28
         X1              =   1710
         X2              =   1710
         Y1              =   3495
         Y2              =   4995
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   26
         X1              =   1710
         X2              =   1710
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   17
         X1              =   1695
         X2              =   1695
         Y1              =   105
         Y2              =   2160
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   16
         X1              =   1710
         X2              =   1710
         Y1              =   105
         Y2              =   2175
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   25
         X1              =   1170
         X2              =   1170
         Y1              =   3495
         Y2              =   4980
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   24
         X1              =   1185
         X2              =   1185
         Y1              =   3495
         Y2              =   4995
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   22
         X1              =   1185
         X2              =   1185
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   5
         X1              =   1170
         X2              =   1170
         Y1              =   105
         Y2              =   2160
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   4
         X1              =   1185
         X2              =   1185
         Y1              =   105
         Y2              =   2175
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   21
         X1              =   645
         X2              =   645
         Y1              =   3495
         Y2              =   4980
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   20
         X1              =   660
         X2              =   660
         Y1              =   3495
         Y2              =   4995
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   18
         X1              =   660
         X2              =   660
         Y1              =   2520
         Y2              =   3150
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   22
         Left            =   1230
         OLEDropMode     =   1  'Manual
         TabIndex        =   135
         Top             =   3225
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   21
         Left            =   1230
         OLEDropMode     =   1  'Manual
         TabIndex        =   134
         Top             =   2250
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   3
         Left            =   705
         OLEDropMode     =   1  'Manual
         TabIndex        =   133
         Top             =   5055
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   2
         Left            =   705
         OLEDropMode     =   1  'Manual
         TabIndex        =   132
         Top             =   3225
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   23
         Left            =   1230
         OLEDropMode     =   1  'Manual
         TabIndex        =   131
         Top             =   5055
         Width           =   435
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   1
         Left            =   705
         OLEDropMode     =   1  'Manual
         TabIndex        =   130
         Top             =   2250
         Width           =   435
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Triple"
         Height          =   195
         Index           =   12
         Left            =   135
         OLEDropMode     =   1  'Manual
         TabIndex        =   127
         Top             =   3855
         Width           =   390
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Yamb"
         Height          =   195
         Index           =   11
         Left            =   128
         OLEDropMode     =   1  'Manual
         TabIndex        =   123
         Top             =   4710
         Width           =   405
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Poker"
         Height          =   195
         Index           =   10
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   122
         Top             =   4425
         Width           =   420
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Full"
         Height          =   195
         Index           =   9
         Left            =   210
         OLEDropMode     =   1  'Manual
         TabIndex        =   121
         Top             =   4140
         Width           =   240
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Straight"
         Height          =   195
         Index           =   8
         Left            =   60
         OLEDropMode     =   1  'Manual
         TabIndex        =   120
         Top             =   3570
         Width           =   540
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Min"
         Height          =   195
         Index           =   7
         Left            =   210
         OLEDropMode     =   1  'Manual
         TabIndex        =   119
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "Max"
         Height          =   195
         Index           =   6
         Left            =   180
         OLEDropMode     =   1  'Manual
         TabIndex        =   118
         Top             =   2610
         Width           =   300
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "6"
         Height          =   195
         Index           =   5
         Left            =   285
         OLEDropMode     =   1  'Manual
         TabIndex        =   117
         Top             =   1905
         Width           =   90
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   195
         Index           =   4
         Left            =   285
         OLEDropMode     =   1  'Manual
         TabIndex        =   116
         Top             =   1620
         Width           =   90
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   195
         Index           =   3
         Left            =   285
         OLEDropMode     =   1  'Manual
         TabIndex        =   115
         Top             =   1335
         Width           =   90
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   195
         Index           =   2
         Left            =   285
         OLEDropMode     =   1  'Manual
         TabIndex        =   114
         Top             =   1050
         Width           =   90
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   195
         Index           =   1
         Left            =   285
         OLEDropMode     =   1  'Manual
         TabIndex        =   113
         Top             =   765
         Width           =   90
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   195
         Index           =   0
         Left            =   285
         OLEDropMode     =   1  'Manual
         TabIndex        =   112
         Top             =   480
         Width           =   90
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000014&
         Index           =   3
         X1              =   660
         X2              =   660
         Y1              =   105
         Y2              =   2175
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   2
         X1              =   645
         X2              =   645
         Y1              =   105
         Y2              =   2160
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   4870
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   7
         X1              =   15
         X2              =   4870
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   11
         X1              =   15
         X2              =   4845
         Y1              =   3135
         Y2              =   3135
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   15
         X1              =   0
         X2              =   4845
         Y1              =   4965
         Y2              =   4965
      End
      Begin VB.Shape SumBackShape 
         BackColor       =   &H80000014&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Index           =   0
         Left            =   15
         Top             =   2175
         Width           =   5370
      End
      Begin VB.Shape SumBackShape 
         BackColor       =   &H80000014&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Index           =   1
         Left            =   30
         Top             =   3150
         Width           =   5355
      End
      Begin VB.Shape SumBackShape 
         BackColor       =   &H80000014&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Index           =   2
         Left            =   15
         Top             =   4980
         Width           =   5370
      End
      Begin VB.Line Line 
         BorderColor     =   &H80000010&
         Index           =   9
         X1              =   0
         X2              =   4850
         Y1              =   2505
         Y2              =   2505
      End
      Begin VB.Shape SumBackShape 
         BackColor       =   &H80000014&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   5190
         Index           =   3
         Left            =   4875
         Top             =   105
         Width           =   510
      End
   End
   Begin VB.Frame Frame 
      Height          =   465
      Index           =   1
      Left            =   0
      OLEDropMode     =   1  'Manual
      TabIndex        =   164
      Top             =   5175
      Width           =   4575
      Begin winYAMB.StatusBox UndoStat 
         Height          =   255
         Left            =   60
         ToolTipText     =   "Undo status"
         Top             =   150
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   450
         Align           =   2
      End
      Begin winYAMB.StatusBox RollStat 
         Height          =   255
         Left            =   375
         Top             =   150
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   450
         Align           =   2
      End
      Begin winYAMB.StatusBox TurnStat 
         Height          =   255
         Left            =   1605
         Top             =   150
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   450
         Align           =   2
      End
      Begin winYAMB.StatusBox MaxStat 
         Height          =   255
         Left            =   2835
         Top             =   150
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   450
         Align           =   2
      End
      Begin winYAMB.StatusBox MinStat 
         Height          =   255
         Left            =   3690
         Top             =   150
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   450
         Align           =   2
      End
   End
   Begin VB.Frame Frame 
      Height          =   2970
      Index           =   3
      Left            =   5415
      OLEDropMode     =   1  'Manual
      TabIndex        =   167
      Top             =   2670
      Width           =   2655
      Begin VB.TextBox NewName 
         Height          =   285
         Index           =   4
         Left            =   390
         OLEDropMode     =   1  'Manual
         TabIndex        =   109
         Top             =   1905
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox NewName 
         Height          =   285
         Index           =   3
         Left            =   390
         OLEDropMode     =   1  'Manual
         TabIndex        =   108
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox NewName 
         Height          =   285
         Index           =   2
         Left            =   390
         OLEDropMode     =   1  'Manual
         TabIndex        =   107
         Top             =   1215
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox NewName 
         Height          =   285
         Index           =   5
         Left            =   390
         OLEDropMode     =   1  'Manual
         TabIndex        =   110
         Top             =   2250
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox NewName 
         Height          =   285
         Index           =   1
         Left            =   390
         OLEDropMode     =   1  'Manual
         TabIndex        =   106
         Top             =   870
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label LabName 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   405
         OLEDropMode     =   1  'Manual
         TabIndex        =   185
         Top             =   900
         Width           =   1410
      End
      Begin VB.Label LabName 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   2
         Left            =   405
         OLEDropMode     =   1  'Manual
         TabIndex        =   184
         Top             =   1245
         Width           =   1410
      End
      Begin VB.Label LabName 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   3
         Left            =   405
         OLEDropMode     =   1  'Manual
         TabIndex        =   183
         Top             =   1590
         Width           =   1410
      End
      Begin VB.Label LabName 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   4
         Left            =   405
         OLEDropMode     =   1  'Manual
         TabIndex        =   182
         Top             =   1935
         Width           =   1410
      End
      Begin VB.Label LabName 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   5
         Left            =   405
         OLEDropMode     =   1  'Manual
         TabIndex        =   181
         Top             =   2280
         Width           =   1410
      End
      Begin VB.Label LabScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   1935
         OLEDropMode     =   1  'Manual
         TabIndex        =   180
         Top             =   900
         Width           =   600
      End
      Begin VB.Label LabScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   2
         Left            =   1935
         OLEDropMode     =   1  'Manual
         TabIndex        =   179
         Top             =   1245
         Width           =   600
      End
      Begin VB.Label LabScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   3
         Left            =   1935
         OLEDropMode     =   1  'Manual
         TabIndex        =   178
         Top             =   1590
         Width           =   600
      End
      Begin VB.Label LabScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   4
         Left            =   1935
         OLEDropMode     =   1  'Manual
         TabIndex        =   177
         Top             =   1935
         Width           =   600
      End
      Begin VB.Label LabScore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   5
         Left            =   1935
         OLEDropMode     =   1  'Manual
         TabIndex        =   176
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   24
         Left            =   405
         OLEDropMode     =   1  'Manual
         TabIndex        =   175
         Top             =   555
         Width           =   420
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score"
         Height          =   195
         Index           =   22
         Left            =   2010
         OLEDropMode     =   1  'Manual
         TabIndex        =   174
         Top             =   555
         Width           =   420
      End
      Begin VB.Line Line 
         BorderColor     =   &H8000000D&
         Index           =   70
         X1              =   75
         X2              =   2580
         Y1              =   825
         Y2              =   825
      End
      Begin VB.Line Line 
         BorderColor     =   &H8000000D&
         Index           =   69
         X1              =   75
         X2              =   2595
         Y1              =   1170
         Y2              =   1170
      End
      Begin VB.Line Line 
         BorderColor     =   &H8000000D&
         Index           =   68
         X1              =   60
         X2              =   2580
         Y1              =   1860
         Y2              =   1860
      End
      Begin VB.Line Line 
         BorderColor     =   &H8000000D&
         Index           =   67
         X1              =   345
         X2              =   345
         Y1              =   495
         Y2              =   2895
      End
      Begin VB.Line Line 
         BorderColor     =   &H8000000D&
         Index           =   66
         X1              =   1860
         X2              =   1860
         Y1              =   495
         Y2              =   2895
      End
      Begin VB.Line Line 
         BorderColor     =   &H8000000D&
         Index           =   14
         X1              =   75
         X2              =   2580
         Y1              =   1515
         Y2              =   1515
      End
      Begin VB.Line Line 
         BorderColor     =   &H8000000D&
         Index           =   10
         X1              =   60
         X2              =   2580
         Y1              =   2205
         Y2              =   2205
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         Height          =   195
         Index           =   21
         Left            =   165
         OLEDropMode     =   1  'Manual
         TabIndex        =   173
         Top             =   2295
         Width           =   90
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         Height          =   195
         Index           =   20
         Left            =   165
         OLEDropMode     =   1  'Manual
         TabIndex        =   172
         Top             =   1935
         Width           =   90
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         Height          =   195
         Index           =   19
         Left            =   165
         OLEDropMode     =   1  'Manual
         TabIndex        =   171
         Top             =   1605
         Width           =   90
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         Height          =   195
         Index           =   18
         Left            =   165
         OLEDropMode     =   1  'Manual
         TabIndex        =   170
         Top             =   1260
         Width           =   90
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         Height          =   195
         Index           =   17
         Left            =   165
         OLEDropMode     =   1  'Manual
         TabIndex        =   169
         Top             =   915
         Width           =   90
      End
      Begin VB.Line Line 
         BorderColor     =   &H8000000D&
         Index           =   6
         X1              =   75
         X2              =   2580
         Y1              =   2550
         Y2              =   2550
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Best Scores"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   16
         Left            =   900
         OLEDropMode     =   1  'Manual
         TabIndex        =   168
         Top             =   225
         Width           =   855
      End
      Begin VB.Shape Shape 
         BackColor       =   &H80000014&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000D&
         Height          =   2415
         Index           =   1
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   2535
      End
      Begin VB.Shape Shape 
         BackColor       =   &H8000000D&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000D&
         Height          =   2235
         Index           =   2
         Left            =   60
         Shape           =   4  'Rounded Rectangle
         Top             =   165
         Width           =   2535
      End
   End
   Begin VB.Frame Frame 
      Height          =   465
      Index           =   4
      Left            =   4590
      OLEDropMode     =   1  'Manual
      TabIndex        =   186
      Top             =   5175
      Width           =   810
      Begin VB.PictureBox SumLabel 
         AutoSize        =   -1  'True
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Height          =   135
         Index           =   4
         Left            =   105
         OLEDropMode     =   1  'Manual
         Picture         =   "Table.frx":10B0
         ScaleHeight     =   135
         ScaleWidth      =   90
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   210
         Width           =   90
      End
      Begin VB.Label SumBox 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   300
         OLEDropMode     =   1  'Manual
         TabIndex        =   188
         Top             =   180
         Width           =   435
      End
      Begin VB.Shape SumBackShape 
         BackColor       =   &H80000014&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   345
         Index           =   4
         Left            =   15
         Top             =   105
         Width           =   780
      End
   End
   Begin VB.Image IcoCnt 
      Height          =   195
      Index           =   7
      Left            =   8160
      Picture         =   "Table.frx":10F1
      Top             =   2640
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image IcoCnt 
      Height          =   195
      Index           =   6
      Left            =   8160
      Picture         =   "Table.frx":1211
      Top             =   2400
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image IcoCnt 
      Height          =   195
      Index           =   5
      Left            =   8160
      Picture         =   "Table.frx":1335
      Top             =   2160
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image IcoCnt 
      Height          =   195
      Index           =   4
      Left            =   8160
      Picture         =   "Table.frx":1431
      Top             =   1920
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image IcoCnt 
      Height          =   195
      Index           =   3
      Left            =   8160
      Picture         =   "Table.frx":1559
      Top             =   1680
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image IcoCnt 
      Height          =   195
      Index           =   2
      Left            =   8160
      Picture         =   "Table.frx":167D
      Top             =   1440
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image IcoCnt 
      Height          =   195
      Index           =   1
      Left            =   8160
      Picture         =   "Table.frx":17A1
      Top             =   1200
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image IcoCnt 
      Height          =   195
      Index           =   0
      Left            =   8160
      Picture         =   "Table.frx":18C1
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image DiceCnt 
      Height          =   240
      Index           =   3
      Left            =   8160
      Picture         =   "Table.frx":19E1
      Top             =   630
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image DiceCnt 
      Height          =   240
      Index           =   2
      Left            =   8145
      Picture         =   "Table.frx":1F6B
      Top             =   360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image DiceCnt 
      Height          =   240
      Index           =   1
      Left            =   8160
      Picture         =   "Table.frx":24F5
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image PictureCnt 
      Height          =   780
      Index           =   6
      Left            =   9120
      Picture         =   "Table.frx":2A7F
      Top             =   120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image PictureCnt 
      Height          =   780
      Index           =   5
      Left            =   9000
      Picture         =   "Table.frx":48D9
      Top             =   120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image PictureCnt 
      Height          =   780
      Index           =   4
      Left            =   8880
      Picture         =   "Table.frx":6733
      Top             =   120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image PictureCnt 
      Height          =   780
      Index           =   3
      Left            =   8760
      Picture         =   "Table.frx":858D
      Top             =   120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image PictureCnt 
      Height          =   780
      Index           =   2
      Left            =   8640
      Picture         =   "Table.frx":A3E7
      Top             =   120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image PictureCnt 
      Height          =   780
      Index           =   1
      Left            =   8520
      Picture         =   "Table.frx":C241
      Top             =   120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNew 
         Caption         =   "&New game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open game..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save game..."
         Shortcut        =   ^S
      End
      Begin VB.Menu Separator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuAppear 
         Caption         =   "&Appearance"
         Begin VB.Menu mnuRect 
            Caption         =   "&Rectangle"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuOval 
            Caption         =   "&Oval"
         End
      End
      Begin VB.Menu mnuBest 
         Caption         =   "&Reset Best Scores"
      End
      Begin VB.Menu Separator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuRules 
         Caption         =   "&Rules..."
      End
      Begin VB.Menu Separator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ClearRolls As Boolean
Private Played As Boolean
Private First As Boolean
Private Second As Boolean
Private Third As Boolean
Private Announced As Boolean
Private Selected As Integer
Private SelectTransfered As Boolean
Private Entered As Integer
Private Finished As Boolean
Private EnableArr As New Collection
Private NewBest As Integer
Private Drop As String
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wflags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Sub CalcSum(Obj As Object, Col As Integer)
Sum1:
    Sum = 0
    For I = Col + 1 To Col + 6
        If Obj(I).Text = "" Then
            GoTo Sum2
        Else
            If IsNumeric(Obj(I).Text) Then
                Sum = Sum + CInt(Obj(I).Text)
            End If
        End If
    Next
    If Sum > 59 Then
        Sum = Sum + 30
    End If
    If Sum = 0 Then
        SumBox(Col + 1).Caption = "/"
    Else
        SumBox(Col + 1).Caption = Sum
    End If
    CalcSumTotal
Sum2:
    Dif = 0
    If Obj(Col + 1).Text <> vbNullString And Obj(Col + 7).Text <> vbNullString And Obj(Col + 8).Text <> vbNullString Then
        If IsNumeric(Obj(Col + 1).Text) Then
            Ones = CInt(Obj(Col + 1).Text)
        Else
            Ones = 0
        End If
        If IsNumeric(Obj(Col + 7).Text) Then
            If IsNumeric(Obj(Col + 8).Text) Then
                If CInt(Obj(Col + 7).Text) > CInt(Obj(Col + 8).Text) Then
                    Dif = (CInt(Obj(Col + 7).Text) - CInt(Obj(Col + 8).Text)) * Ones
                Else
                    Dif = 0
                End If
            End If
        End If
        If Dif = 0 Then
            SumBox(Col + 2).Caption = "/"
        Else
            SumBox(Col + 2).Caption = Dif
        End If
    End If
    CalcSumTotal
Sum3:
    Sum = 0
    For I = Col + 9 To Col + 13
        If Obj(I).Text = "" Then
            Exit Sub
        Else
            If IsNumeric(Obj(I).Text) Then
                Sum = Sum + CInt(Obj(I).Text)
            End If
        End If
    Next
    If Sum = 0 Then
        SumBox(Col + 3).Caption = "/"
    Else
        SumBox(Col + 3).Caption = Sum
    End If
    CalcSumTotal
End Sub
Private Sub CalcSumTotal()
    For I = 1 To 3
        If SumBox(I).Caption <> vbNullString And _
           SumBox(20 + I).Caption <> vbNullString And _
           SumBox(40 + I).Caption <> vbNullString And _
           SumBox(60 + I).Caption <> vbNullString And _
           SumBox(80 + I).Caption <> vbNullString And _
           SumBox(100 + I).Caption <> vbNullString And _
           SumBox(120 + I).Caption <> vbNullString And _
           SumBox(140 + I).Caption <> vbNullString Then
            Sum = 0
            If IsNumeric(SumBox(I).Caption) Then Sum = Sum + CInt(SumBox(I).Caption)
            If IsNumeric(SumBox(20 + I).Caption) Then Sum = Sum + CInt(SumBox(20 + I).Caption)
            If IsNumeric(SumBox(40 + I).Caption) Then Sum = Sum + CInt(SumBox(40 + I).Caption)
            If IsNumeric(SumBox(60 + I).Caption) Then Sum = Sum + CInt(SumBox(60 + I).Caption)
            If IsNumeric(SumBox(80 + I).Caption) Then Sum = Sum + CInt(SumBox(80 + I).Caption)
            If IsNumeric(SumBox(100 + I).Caption) Then Sum = Sum + CInt(SumBox(100 + I).Caption)
            If IsNumeric(SumBox(120 + I).Caption) Then Sum = Sum + CInt(SumBox(120 + I).Caption)
            If IsNumeric(SumBox(140 + I).Caption) Then Sum = Sum + CInt(SumBox(140 + I).Caption)
            If Sum = 0 Then
                SumBox(160 + I).Caption = "/"
            Else
                SumBox(160 + I).Caption = Sum
            End If
        End If
    Next
    If SumBox(161).Caption <> vbNullString And _
       SumBox(162).Caption <> vbNullString And _
       SumBox(163).Caption <> vbNullString Then
        Sum = 0
        If IsNumeric(SumBox(161).Caption) Then Sum = Sum + CInt(SumBox(161).Caption)
        If IsNumeric(SumBox(162).Caption) Then Sum = Sum + CInt(SumBox(162).Caption)
        If IsNumeric(SumBox(163).Caption) Then Sum = Sum + CInt(SumBox(163).Caption)
        If Sum = 0 Then
            SumBox(0).Caption = "/"
        Else
            SumBox(0).Caption = Sum
        End If
        Finished = True
    End If
End Sub
Private Function CalcSame$(Num)
    CalcSame = 0
    TArr = Array(0, 0, 0, 0, 0, 0)
    For I = 1 To 6
        Select Case Dice(I).Tag
            Case 1
                TArr(0) = TArr(0) + 1
            Case 2
                TArr(1) = TArr(1) + 1
            Case 3
                TArr(2) = TArr(2) + 1
            Case 4
                TArr(3) = TArr(3) + 1
            Case 5
                TArr(4) = TArr(4) + 1
            Case 6
                TArr(5) = TArr(5) + 1
        End Select
    Next
    For I = 5 To 0 Step -1
        If TArr(I) > Num - 1 Then
            Select Case Num
                Case 3
                    CalcSame = (I + 1) * Num + 20
                Case 4 To 5
                    CalcSame = (I + 1) * Num + Num * 10
            End Select
            Exit For
        End If
    Next
End Function
Private Function CalcField$(Idx As Integer)
    Select Case Idx
        Case 1 To 13
            Field = Idx
        Case 21 To 33
            Field = Idx - 20
        Case 41 To 53
            Field = Idx - 40
        Case 61 To 73
            Field = Idx - 60
        Case 81 To 93
            Field = Idx - 80
        Case 101 To 113
            Field = Idx - 100
        Case 121 To 133
            Field = Idx - 120
        Case 141 To 153
            Field = Idx - 140
    End Select
    Select Case Field
        Case 0 To 6
            Value = 0
            For I = 1 To 6
                If CInt(Dice(I).Tag) = Field Then
                    Value = Value + CInt(Dice(I).Tag)
                End If
            Next
            If Value > Field * 5 Then Value = Field * 5
            CalcField = Value
        Case 7 To 8
            Value = 0
            If Field = 7 Then
                A = 1
                B = 6
                c = 1
            Else
                A = 6
                B = 1
                c = -1
            End If
            For I = A To B Step c
                For iter = 1 To 6
                    If CInt(Dice(iter).Tag) = I Then
                        Subtraction = CInt(Dice(iter).Tag)
                        Exit For
                    End If
                Next
                If Subtraction <> 0 Then Exit For
            Next
            For I = 1 To 6
                Value = Value + CInt(Dice(I).Tag)
            Next
            CalcField = Value - Subtraction
        Case 9
            Value = 0
            For I = 1 To 6
                Select Case Dice(I).Tag
                    Case 1
                        One = True
                    Case 2
                        Two = True
                    Case 3
                        Three = True
                    Case 4
                        Four = True
                    Case 5
                        Five = True
                    Case 6
                        Six = True
                End Select
            Next
            If One And Two And Three And Four And Five Or _
               Two And Three And Four And Five And Six Then
                If First Then Value = 66
                If Second Then Value = 56
                If Third Then Value = 46
            End If
            CalcField = Value
        Case 10
            Value = CalcSame(3)
            CalcField = Value
        Case 11
            Value = 0
            Cond = False
            TArr = Array(0, 0, 0, 0, 0, 0)
            For I = 1 To 6
                Select Case Dice(I).Tag
                    Case 1
                        TArr(0) = TArr(0) + 1
                    Case 2
                        TArr(1) = TArr(1) + 1
                    Case 3
                        TArr(2) = TArr(2) + 1
                    Case 4
                        TArr(3) = TArr(3) + 1
                    Case 5
                        TArr(4) = TArr(4) + 1
                    Case 6
                        TArr(5) = TArr(5) + 1
                End Select
            Next
            For I = 5 To 0 Step -1
                For iter = 5 To 0 Step -1
                    If iter <> I Then
                        If TArr(I) > 2 And TArr(iter) > 1 Then
                            Cond = True
                            Major = I + 1
                            Minor = iter + 1
                        End If
                    End If
                    If Cond Then Exit For
                Next
                If Cond Then Exit For
            Next
            If Cond Then
                Value = 3 * Major + 2 * Minor + 30
            End If
            CalcField = Value
        Case 12
            Value = CalcSame(4)
            CalcField = Value
        Case 13
            Value = CalcSame(5)
            CalcField = Value
    End Select
    If Not (Announced And Second) Then mnuUndo.Enabled = True
    RefreshStatus
End Function
Private Sub ColLabel_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub Down_Click(Index As Integer)
    If First Then
        Value = CalcField(Index)
        If Value = 0 Then
            Down(Index).Text = "/"
        Else
            Down(Index).Text = Value
        End If
        Entered = Index
        Disable
        CalcSum Down, 0
        TurnBtn.Enabled = True
        TurnBtn.SetFocus
    Else
        MsgBox "Roll dices first!"
    End If
End Sub
Private Sub Down_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace, vbKeyAdd
            Down_Click Index
    End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyNumpad4
            Dice_Click 1
        Case vbKeyNumpad5
            Dice_Click 2
        Case vbKeyNumpad6
            Dice_Click 3
        Case vbKeyNumpad1
            Dice_Click 4
        Case vbKeyNumpad2
            Dice_Click 5
        Case vbKeyNumpad3
            Dice_Click 6
        Case vbKeySubtract
            keybd_event &H9, MapVirtualKey(&H9, 0), 0, 0
            keybd_event &H9, MapVirtualKey(&H9, 0), &H2 Or 0, 0
        Case vbKeyDivide
            keybd_event &H10, MapVirtualKey(&H10, 0), 0, 0
            keybd_event &H9, MapVirtualKey(&H9, 0), 0, 0
            keybd_event &H9, MapVirtualKey(&H9, 0), &H2 Or 0, 0
            keybd_event &H10, MapVirtualKey(&H10, 0), &H2 Or 0, 0
        Case vbKeyMultiply
            If RollBtn.Enabled Then RollBtn_Click
        Case vbKeyNumpad0
            If TurnBtn.Enabled Then TurnBtn_Click
    End Select
End Sub
Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Drop = Data.Files.Item(1)
    mnuOpen_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
Private Sub Frame_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub IcoLabel_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub Label_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub LabHit_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub LabName_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub LabScore_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub mnuAbout_Click()
    Splash.Show
End Sub
Private Sub mnuBest_Click()
    If MsgBox("Are you sure you want to reset best scores?", vbYesNo + vbQuestion) = vbYes Then
        If MsgBox("Are you REALLY sure you want to reset best scores?" & Chr(10) & "This operation is not undoable.", vbYesNo + vbQuestion) = vbYes Then
            Set fso = CreateObject("Scripting.FileSystemObject")
            If fso.FileExists("winYAMB.dat") Then
                fso.DeleteFile "winYAMB.dat", True
            End If
            LoadBestScores
        End If
    End If
End Sub
Private Sub mnuContents_Click()
    Help Me, "overview.html"
End Sub
Private Sub mnuExit_Click()
    If MsgBox("Are you sure you want to quit this game?", vbYesNo + vbQuestion) = vbYes Then Unload Me
End Sub
Private Sub mnuOpen_Click()
    If Finished Or Not Played Or Len(Drop) > 0 Then
        State = True
    Else
        If MsgBox("If you open a game it will replace the game in progress." & Chr(10) & _
                  "Are you sure you want to continue?", vbYesNo + vbQuestion) = vbYes Then State = True
    End If
    If State Then
        If Len(Drop) = 0 Then
            FName = SelectFile(Me.hwnd, "winYAMB File (*.wyb)|*.wyb", "*.wyb", cdfmOpenFile, "Open game...")
        Else
            FName = Drop
            Drop = vbNullString
        End If
        If FName <> vbNullString Then
            OpenIt FName
        End If
    End If
End Sub
Public Sub OpenIt(FName)
    InitFields
    Finished = False
    KillTimer Me.hwnd, 1
    RollBtn.Enabled = True
    LabHit.Visible = False
    For I = 1 To 6
        Dice(I).Visible = True
    Next
    mnuNew.Enabled = True
    mnuSave.Enabled = True
    mnuBest.Enabled = True
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set GameFile = fso.OpenTextFile(FName, 1)
    GameStr = vbNullString
    Do Until GameFile.AtEndOfStream
        GameStr = GameStr & GameFile.ReadLine
    Loop
    GameFile.Close
    DecStr = DecryptText(GameStr, "hJs6OPd")
    GameArr = Split(DecStr, "#DATA_EDGE#", -1, vbTextCompare)
    TableArr = Split(GameArr(0), "#NEW_COLUMN#", -1, vbTextCompare)
    For I = 0 To 7
        ColArr = Split(TableArr(I), "#NEW_FIELD#", -1, vbTextCompare)
        Select Case I
            Case 0
                For iter = 0 To 12
                    Down(iter + 1).Text = ColArr(iter)
                    Down(iter + 1).Enabled = False
                Next
            Case 1
                For iter = 0 To 12
                    Random(iter + 21).Text = ColArr(iter)
                    Random(iter + 21).Enabled = False
                Next
            Case 2
                For iter = 0 To 12
                    Up(iter + 41).Text = ColArr(iter)
                    Up(iter + 41).Enabled = False
                Next
            Case 3
                For iter = 0 To 12
                    Announce(iter + 61).Text = ColArr(iter)
                    Announce(iter + 61).Enabled = False
                Next
            Case 4
                For iter = 0 To 12
                    Hand(iter + 81).Text = ColArr(iter)
                    Hand(iter + 81).Enabled = False
                Next
            Case 5
                For iter = 0 To 12
                    Middle(iter + 101).Text = ColArr(iter)
                    Middle(iter + 101).Enabled = False
                Next
            Case 6
                For iter = 0 To 12
                    UpDown(iter + 121).Text = ColArr(iter)
                    UpDown(iter + 121).Enabled = False
                Next
            Case 7
                For iter = 0 To 12
                    Max(iter + 141).Text = ColArr(iter)
                    Max(iter + 141).Enabled = False
                Next
        End Select
    Next
    SumsArr = Split(GameArr(1), "#NEXT_COLUMN#", -1, vbTextCompare)
    For I = 0 To 8
        SumArr = Split(SumsArr(I), "#NEXT_SUM#", -1, vbTextCompare)
        For iter = 0 To 2
            SumBox(iter + (I * 20) + 1).Caption = SumArr(iter)
        Next
    Next
    StatesArr = Split(GameArr(2), "#NEW_COLUMN#", -1, vbTextCompare)
    For I = 0 To 7
        StateArr = Split(StatesArr(I), "#NEW_STATE#", -1, vbTextCompare)
        Select Case I
            Case 0
                For iter = 0 To 12
                    Down(iter + 1).Enabled = CBool(StateArr(iter))
                Next
            Case 1
                For iter = 0 To 12
                    Random(iter + 21).Enabled = CBool(StateArr(iter))
                Next
            Case 2
                For iter = 0 To 12
                    Up(iter + 41).Enabled = CBool(StateArr(iter))
                Next
            Case 3
                For iter = 0 To 12
                    Announce(iter + 61).Enabled = CBool(StateArr(iter))
                Next
            Case 4
                For iter = 0 To 12
                    Hand(iter + 81).Enabled = CBool(StateArr(iter))
                Next
            Case 5
                For iter = 0 To 12
                    Middle(iter + 101).Enabled = CBool(StateArr(iter))
                Next
            Case 6
                For iter = 0 To 12
                    UpDown(iter + 121).Enabled = CBool(StateArr(iter))
                Next
            Case 7
                For iter = 0 To 12
                    Max(iter + 141).Enabled = CBool(StateArr(iter))
                Next
        End Select
    Next
    SlashPos = InStrRev(FName, "\")
    FName = Right(FName, Len(FName) - SlashPos)
    Table.Caption = "winYAMB - " & FName
    RefreshStatus
End Sub
Public Sub mnuOval_Click()
    If Not mnuOval.Checked Then
        mnuOval.Checked = True
        mnuRect.Checked = False
        For I = 1 To 153
            Select Case I
                Case 1 To 13
                    Down(I).Shape = 2
                Case 21 To 33
                    Random(I).Shape = 2
                Case 41 To 53
                    Up(I).Shape = 2
                Case 61 To 73
                    Announce(I).Shape = 2
                Case 81 To 93
                    Hand(I).Shape = 2
                Case 101 To 113
                    Middle(I).Shape = 2
                Case 121 To 133
                    UpDown(I).Shape = 2
                Case 141 To 153
                    Max(I).Shape = 2
            End Select
        Next
        regCreate_Key_Value HKEY_LOCAL_MACHINE, "Software\.dUcA\winYAMB", "Appearance", "@"
    End If
End Sub
Private Sub mnuRect_Click()
    If Not mnuRect.Checked Then
        mnuRect.Checked = True
        mnuOval.Checked = False
        For I = 1 To 153
            Select Case I
                Case 1 To 13
                    Down(I).Shape = 0
                Case 21 To 33
                    Random(I).Shape = 0
                Case 41 To 53
                    Up(I).Shape = 0
                Case 61 To 73
                    Announce(I).Shape = 0
                Case 81 To 93
                    Hand(I).Shape = 0
                Case 101 To 113
                    Middle(I).Shape = 0
                Case 121 To 133
                    UpDown(I).Shape = 0
                Case 141 To 153
                    Max(I).Shape = 0
            End Select
        Next
        regDelete_Sub_Key HKEY_LOCAL_MACHINE, "Software\.dUcA\winYAMB", "Appearance"
    End If
End Sub
Private Sub mnuRules_Click()
    Help Me, "rules.html"
End Sub
Private Sub mnuSave_Click()
    If Not First Then
        FName = SelectFile(Me.hwnd, "winYAMB File (*.wyb)|*.wyb", "*.wyb", cdfmSaveFileNoConfirm, "Save game...", "Game1.wyb")
        If FName <> vbNullString Then
            SaveStr = vbNullString
            For I = 1 To 153
                Select Case I
                    Case 1 To 12
                        SaveStr = SaveStr & Down(I).Text & "#NEW_FIELD#"
                    Case 13
                        SaveStr = SaveStr & Down(I).Text & "#NEW_COLUMN#"
                    Case 21 To 32
                        SaveStr = SaveStr & Random(I).Text & "#NEW_FIELD#"
                    Case 33
                        SaveStr = SaveStr & Random(I).Text & "#NEW_COLUMN#"
                    Case 41 To 52
                        SaveStr = SaveStr & Up(I).Text & "#NEW_FIELD#"
                    Case 53
                        SaveStr = SaveStr & Up(I).Text & "#NEW_COLUMN#"
                    Case 61 To 72
                        SaveStr = SaveStr & Announce(I).Text & "#NEW_FIELD#"
                    Case 73
                        SaveStr = SaveStr & Announce(I).Text & "#NEW_COLUMN#"
                    Case 81 To 92
                        SaveStr = SaveStr & Hand(I).Text & "#NEW_FIELD#"
                    Case 93
                        SaveStr = SaveStr & Hand(I).Text & "#NEW_COLUMN#"
                    Case 101 To 112
                        SaveStr = SaveStr & Middle(I).Text & "#NEW_FIELD#"
                    Case 113
                        SaveStr = SaveStr & Middle(I).Text & "#NEW_COLUMN#"
                    Case 121 To 132
                        SaveStr = SaveStr & UpDown(I).Text & "#NEW_FIELD#"
                    Case 133
                        SaveStr = SaveStr & UpDown(I).Text & "#NEW_COLUMN#"
                    Case 141 To 152
                        SaveStr = SaveStr & Max(I).Text & "#NEW_FIELD#"
                    Case 153
                        SaveStr = SaveStr & Max(I).Text & "#DATA_EDGE#"
                End Select
            Next
            For I = 0 To 8
                SaveStr = SaveStr & SumBox((I * 20) + 1).Caption & "#NEXT_SUM#"
                SaveStr = SaveStr & SumBox((I * 20) + 2).Caption & "#NEXT_SUM#"
                SaveStr = SaveStr & SumBox((I * 20) + 3).Caption & "#NEXT_COLUMN#"
            Next
            SaveStr = Left(SaveStr, Len(SaveStr) - 13) & "#DATA_EDGE#"
            For I = 1 To 153
                Select Case I
                    Case 1 To 12
                        SaveStr = SaveStr & CInt(Down(I).Enabled) & "#NEW_STATE#"
                    Case 13
                        SaveStr = SaveStr & CInt(Down(I).Enabled) & "#NEW_COLUMN#"
                    Case 21 To 32
                        SaveStr = SaveStr & CInt(Random(I).Enabled) & "#NEW_STATE#"
                    Case 33
                        SaveStr = SaveStr & CInt(Random(I).Enabled) & "#NEW_COLUMN#"
                    Case 41 To 52
                        SaveStr = SaveStr & CInt(Up(I).Enabled) & "#NEW_STATE#"
                    Case 53
                        SaveStr = SaveStr & CInt(Up(I).Enabled) & "#NEW_COLUMN#"
                    Case 61 To 72
                        SaveStr = SaveStr & CInt(Announce(I).Enabled) & "#NEW_STATE#"
                    Case 73
                        SaveStr = SaveStr & CInt(Announce(I).Enabled) & "#NEW_COLUMN#"
                    Case 81 To 92
                        SaveStr = SaveStr & CInt(Hand(I).Enabled) & "#NEW_STATE#"
                    Case 93
                        SaveStr = SaveStr & CInt(Hand(I).Enabled) & "#NEW_COLUMN#"
                    Case 101 To 112
                        SaveStr = SaveStr & CInt(Middle(I).Enabled) & "#NEW_STATE#"
                    Case 113
                        SaveStr = SaveStr & CInt(Middle(I).Enabled) & "#NEW_COLUMN#"
                    Case 121 To 132
                        SaveStr = SaveStr & CInt(UpDown(I).Enabled) & "#NEW_STATE#"
                    Case 133
                        SaveStr = SaveStr & CInt(UpDown(I).Enabled) & "#NEW_COLUMN#"
                    Case 141 To 152
                        SaveStr = SaveStr & CInt(Max(I).Enabled) & "#NEW_STATE#"
                    Case 153
                        SaveStr = SaveStr & CInt(Max(I).Enabled)
                End Select
            Next
            EncStr = EncryptText(SaveStr, "hJs6OPd")
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set GameFile = fso.OpenTextFile(FName, 2, True)
            GameFile.Write EncStr
            GameFile.Close
            SlashPos = InStrRev(FName, "\")
            FName = Right(FName, Len(FName) - SlashPos)
            Table.Caption = "winYAMB - " & FName
        End If
        Table.SetFocus
    Else
        MsgBox "You can not save game while still in turn." & Chr(10) & _
               "Finish this turn and try again."
    End If
End Sub
Private Sub mnuUndo_Click()
    If First And Entered <> 0 Then
        Select Case Entered
            Case 1 To 6
                Down(Entered).Text = vbNullString
                SumBox(1).Caption = vbNullString
                SumBox(161).Caption = vbNullString
            Case 7 To 8
                Down(Entered).Text = vbNullString
                SumBox(2).Caption = vbNullString
                SumBox(162).Caption = vbNullString
            Case 9 To 13
                Down(Entered).Text = vbNullString
                SumBox(3).Caption = vbNullString
                SumBox(163).Caption = vbNullString
            Case 21 To 26
                Random(Entered).Text = vbNullString
                SumBox(21).Caption = vbNullString
                SumBox(161).Caption = vbNullString
            Case 27 To 28
                Random(Entered).Text = vbNullString
                SumBox(22).Caption = vbNullString
                SumBox(162).Caption = vbNullString
            Case 29 To 33
                Random(Entered).Text = vbNullString
                SumBox(23).Caption = vbNullString
                SumBox(163).Caption = vbNullString
            Case 41 To 46
                Up(Entered).Text = vbNullString
                SumBox(41).Caption = vbNullString
                SumBox(161).Caption = vbNullString
            Case 47 To 48
                Up(Entered).Text = vbNullString
                SumBox(42).Caption = vbNullString
                SumBox(162).Caption = vbNullString
            Case 49 To 53
                Up(Entered).Text = vbNullString
                SumBox(43).Caption = vbNullString
                SumBox(163).Caption = vbNullString
            Case 61 To 66
                Announce(Entered).Text = vbNullString
                Announced = False
                SumBox(61).Caption = vbNullString
                SumBox(161).Caption = vbNullString
            Case 67 To 68
                Announce(Entered).Text = vbNullString
                Announced = False
                SumBox(62).Caption = vbNullString
                SumBox(162).Caption = vbNullString
            Case 69 To 73
                Announce(Entered).Text = vbNullString
                Announced = False
                SumBox(63).Caption = vbNullString
                SumBox(163).Caption = vbNullString
            Case 81 To 93
                GoTo Msg
            Case 101 To 106
                Middle(Entered).Text = vbNullString
                SumBox(101).Caption = vbNullString
                SumBox(161).Caption = vbNullString
            Case 107 To 108
                Middle(Entered).Text = vbNullString
                SumBox(102).Caption = vbNullString
                SumBox(162).Caption = vbNullString
            Case 109 To 113
                Middle(Entered).Text = vbNullString
                SumBox(103).Caption = vbNullString
                SumBox(163).Caption = vbNullString
            Case 121 To 126
                UpDown(Entered).Text = vbNullString
                SumBox(121).Caption = vbNullString
                SumBox(161).Caption = vbNullString
            Case 127 To 128
                UpDown(Entered).Text = vbNullString
                SumBox(122).Caption = vbNullString
                SumBox(162).Caption = vbNullString
            Case 129 To 133
                UpDown(Entered).Text = vbNullString
                SumBox(123).Caption = vbNullString
                SumBox(163).Caption = vbNullString
            Case 141 To 146
                Max(Entered).Text = vbNullString
                SumBox(141).Caption = vbNullString
                SumBox(161).Caption = vbNullString
            Case 147 To 148
                Max(Entered).Text = vbNullString
                SumBox(142).Caption = vbNullString
                SumBox(162).Caption = vbNullString
            Case 149 To 153
                Max(Entered).Text = vbNullString
                SumBox(143).Caption = vbNullString
                SumBox(163).Caption = vbNullString
        End Select
        SumBox(0).Caption = vbNullString
        EnableSelective
        Entered = 0
        TurnBtn.Enabled = False
        RefreshStatus
        mnuUndo.Enabled = False
    Else
Msg:
        MsgBox "Undo option is unavailable between turns" & Chr(10) & _
               "or if no field has been entered.", vbMsgBoxHelpButton, , "winYAMB.chm::/rules.html#undo", 0
    End If
End Sub
Private Sub NewName_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub Random_Click(Index As Integer)
    If First Then
        Value = CalcField(Index)
        If Value = 0 Then
            Random(Index).Text = "/"
        Else
            Random(Index).Text = Value
        End If
        Entered = Index
        Disable
        CalcSum Random, 20
        TurnBtn.Enabled = True
        TurnBtn.SetFocus
    Else
        MsgBox "Roll dices first!"
    End If
End Sub
Private Sub Random_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace, vbKeyAdd
            Random_Click Index
    End Select
End Sub
Private Sub SumBox_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub SumLabel_OLEDragDrop(Index As Integer, Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub
Private Sub Up_Click(Index As Integer)
    If First Then
        Value = CalcField(Index)
        If Value = 0 Then
            Up(Index).Text = "/"
        Else
            Up(Index).Text = Value
        End If
        Entered = Index
        Disable
        CalcSum Up, 40
        TurnBtn.Enabled = True
        TurnBtn.SetFocus
    Else
        MsgBox "Roll dices first!"
    End If
End Sub
Private Sub Up_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace, vbKeyAdd
            Up_Click Index
    End Select
End Sub
Private Sub Announce_Click(Index As Integer)
    If First Then
        If Not Announced Then
            Announced = True
        End If
        Value = 0
        Entered = Index
        Disable
        Value = CalcField(Index)
        If Value = 0 Then
            Announce(Index).Text = "/"
        Else
            Announce(Index).Text = Value
        End If
        CalcSum Announce, 60
        TurnBtn.Enabled = True
        TurnBtn.SetFocus
    Else
        MsgBox "Roll dices first!"
    End If
End Sub
Private Sub Announce_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace, vbKeyAdd
            Announce_Click Index
    End Select
End Sub
Private Sub Hand_Click(Index As Integer)
    If First Then
        Value = CalcField(Index)
        If Value = 0 Then
            Hand(Index).Text = "/"
        Else
            Hand(Index).Text = Value
        End If
        Entered = Index
        CalcSum Hand, 80
        TurnBtn.Enabled = True
        TurnBtn_Click
    Else
        MsgBox "Roll dices first!"
    End If
End Sub
Private Sub Hand_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace, vbKeyAdd
            Hand_Click Index
    End Select
End Sub
Private Sub Middle_Click(Index As Integer)
    If First Then
        Value = CalcField(Index)
        If Value = 0 Then
            Middle(Index).Text = "/"
        Else
            Middle(Index).Text = Value
        End If
        Entered = Index
        Disable
        CalcSum Middle, 100
        TurnBtn.Enabled = True
        TurnBtn.SetFocus
    Else
        MsgBox "Roll dices first!"
    End If
End Sub
Private Sub Middle_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace, vbKeyAdd
            Middle_Click Index
    End Select
End Sub
Private Sub UpDown_Click(Index As Integer)
    If First Then
        Value = CalcField(Index)
        If Value = 0 Then
            UpDown(Index).Text = "/"
        Else
            UpDown(Index).Text = Value
        End If
        Entered = Index
        Disable
        CalcSum UpDown, 120
        TurnBtn.Enabled = True
        
    Else
        MsgBox "Roll dices first!"
    End If
End Sub
Private Sub UpDown_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace, vbKeyAdd
            UpDown_Click Index
    End Select
End Sub
Private Sub Max_Click(Index As Integer)
    If First Then
        Value = CalcField(Index)
        Select Case Index
            Case 141 To 146
                If Value = 4 * (Index - 140) Or Value = 5 * (Index - 140) Then
                    Max(Index).Text = Value
                Else
                    Max(Index).Text = "/"
                End If
            Case 147
                If Value = 29 Or Value = 30 Then
                    Max(Index).Text = Value
                Else
                    Max(Index).Text = "/"
                End If
            Case 148
                If Value = 5 Or Value = 6 Then
                    Max(Index).Text = Value
                Else
                    Max(Index).Text = "/"
                End If
            Case 149
                If Value = 66 Then
                    Max(Index).Text = Value
                Else
                    Max(Index).Text = "/"
                End If
            Case 150
                If Value = 38 Then
                    Max(Index).Text = Value
                Else
                    Max(Index).Text = "/"
                End If
            Case 151
                If Value = 58 Then
                    Max(Index).Text = Value
                Else
                    Max(Index).Text = "/"
                End If
            Case 152
                If Value = 64 Then
                    Max(Index).Text = Value
                Else
                    Max(Index).Text = "/"
                End If
            Case 153
                If Value = 80 Then
                    Max(Index).Text = Value
                Else
                    Max(Index).Text = "/"
                End If
        End Select
        Entered = Index
        Disable
        CalcSum Max, 140
        TurnBtn.Enabled = True
        TurnBtn.SetFocus
    Else
        MsgBox "Roll dices first!"
    End If
End Sub
Private Sub Max_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace, vbKeyAdd
            Max_Click Index
    End Select
End Sub
Public Sub InitFields()
    DoEvents
    For I = 1 To 13
        Down(I).Text = vbNullString
        Down(I).Enabled = False
        Random(I + 20).Text = vbNullString
        Random(I + 20).Enabled = True
        Up(I + 40).Text = vbNullString
        Up(I + 40).Enabled = False
        Announce(I + 60).Text = vbNullString
        Announce(I + 60).Enabled = True
        Hand(I + 80).Text = vbNullString
        Hand(I + 80).Enabled = True
        Middle(I + 100).Text = vbNullString
        Middle(I + 100).Enabled = False
        UpDown(I + 120).Text = vbNullString
        UpDown(I + 120).Enabled = False
        Max(I + 140).Text = vbNullString
        Max(I + 140).Enabled = True
    Next
    Down(1).Enabled = True
    Up(53).Enabled = True
    Middle(107).Enabled = True
    Middle(108).Enabled = True
    UpDown(121).Enabled = True
    UpDown(133).Enabled = True
    For I = 1 To 3
        SumBox(I).Caption = vbNullString
        SumBox(I + 20).Caption = vbNullString
        SumBox(I + 40).Caption = vbNullString
        SumBox(I + 60).Caption = vbNullString
        SumBox(I + 80).Caption = vbNullString
        SumBox(I + 100).Caption = vbNullString
        SumBox(I + 120).Caption = vbNullString
        SumBox(I + 140).Caption = vbNullString
        SumBox(I + 160).Caption = vbNullString
    Next
    SumBox(0).Caption = vbNullString
    mnuNew.Enabled = False
    mnuUndo.Enabled = False
    mnuSave.Enabled = False
    Played = False
    Finished = False
    Entered = 0
    TurnBtn_Click
    For I = 1 To 5
        NewName(I).Visible = False
    Next
    LoadBestScores
    Me.Refresh
End Sub
Private Sub Form_Load()
    Selected = 0
    Entered = 0
    Set IcoLabel.Picture = DiceCnt(3).Picture
End Sub
Private Sub mnuNew_Click()
    State = False
    If Finished Or Not Played Then
        State = True
    Else
        If MsgBox("Are you sure you want to start new game?", vbYesNo + vbQuestion) = vbYes Then State = True
    End If
    If State Then
        InitFields
        KillTimer Me.hwnd, 1
        LabHit.Visible = False
        For I = 1 To 6
            Dice(I).Visible = True
        Next
        mnuNew.Enabled = False
        mnuSave.Enabled = True
        mnuBest.Enabled = True
        Table.Caption = "winYAMB"
    End If
End Sub
Private Sub Disable()
    On Error Resume Next
    For I = 1 To 153
        Select Case I
            Case 1 To 13
                If Down(I).Enabled And Entered <> I Then
                    EnableArr.Add I, CStr(I)
                    Down(I).Enabled = False
                End If
            Case 21 To 33
                If Random(I).Enabled And Entered <> I Then
                    EnableArr.Add I, CStr(I)
                    Random(I).Enabled = False
                End If
            Case 41 To 53
                If Up(I).Enabled And Entered <> I Then
                    EnableArr.Add I, CStr(I)
                    Up(I).Enabled = False
                End If
            Case 61 To 73
                If Announce(I).Enabled And Entered <> I Then
                    EnableArr.Add I, CStr(I)
                    Announce(I).Enabled = False
                End If
            Case 81 To 93
                If Hand(I).Enabled And Entered <> I Then
                    EnableArr.Add I, CStr(I)
                    Hand(I).Enabled = False
                End If
            Case 101 To 113
                If Middle(I).Enabled And Entered <> I Then
                    EnableArr.Add I, CStr(I)
                    Middle(I).Enabled = False
                End If
            Case 121 To 133
                If UpDown(I).Enabled And Entered <> I Then
                    EnableArr.Add I, CStr(I)
                    UpDown(I).Enabled = False
                End If
            Case 141 To 153
                If Max(I).Enabled And Entered <> I Then
                    EnableArr.Add I, CStr(I)
                    Max(I).Enabled = False
                End If
        End Select
    Next
    RefreshStatus
End Sub
Private Sub Enable()
    For Each Enabler In EnableArr
        If Not Enabler = Entered Then
        Select Case Enabler
            Case 1 To 13
                Down(Enabler).Enabled = True
            Case 21 To 33
                Random(Enabler).Enabled = True
            Case 41 To 53
                Up(Enabler).Enabled = True
            Case 61 To 73
                Announce(Enabler).Enabled = True
            Case 81 To 93
                Hand(Enabler).Enabled = True
            Case 101 To 113
                Middle(Enabler).Enabled = True
            Case 121 To 133
                UpDown(Enabler).Enabled = True
            Case 141 To 153
                Max(Enabler).Enabled = True
        End Select
        End If
    Next
    For I = 1 To EnableArr.Count
        EnableArr.Remove 1
    Next
End Sub
Private Sub EnableSelective()
    For Each Enabler In EnableArr
        Select Case Enabler
            Case Entered
            Case 1 To 13
                Down(Enabler).Enabled = True
            Case 21 To 33
                Random(Enabler).Enabled = True
            Case 41 To 53
                Up(Enabler).Enabled = True
            Case 61 To 73
                If Not Announced And Not Second Then Announce(Enabler).Enabled = True
            Case 81 To 93
                If Not SelectTransfered Then Hand(Enabler).Enabled = True
            Case 101 To 113
                Middle(Enabler).Enabled = True
            Case 121 To 133
                UpDown(Enabler).Enabled = True
            Case 141 To 153
                Max(Enabler).Enabled = True
        End Select
    Next
End Sub
Private Sub RollBtn_Click()
If Not First Then
    If Not Played Then
        Played = True
        mnuNew.Enabled = True
        mnuSave.Enabled = True
    End If
    RndDices
    First = True
    Set RollBtn.Image = DiceCnt(2).Picture
Else
    If ChkAvailable Or Announced Or Entered > 0 Then
        If Announced Then mnuUndo.Enabled = False
        RollType
        If Selected <> 0 Then
            For I = 81 To 93
                If Hand(I).Enabled Then
                    EnableArr.Add I
                    Hand(I).Enabled = False
                End If
            Next
            SelectTransfered = True
        End If
    Else
        If ChkHand Then
            If Selected <> 0 Then
                If ChkAnnounce Then
                    MsgBox "You must select Announce or Hand field."
                Else
                    MsgBox "You must select Hand field."
                End If
            Else
                RollType
            End If
        Else
            If ChkAnnounce Then
                MsgBox "You must select Announce field."
            End If
        End If
    End If
End If
RefreshStatus
End Sub
Private Sub RollType()
    RndDices
    If Second Then
        Third = True
    Else
        Second = True
        Set RollBtn.Image = DiceCnt(1).Picture
    End If
    AutoClick
    If Third Then
        RollBtn.Enabled = False
        Set RollBtn.Image = Nothing
    End If
    RefreshStatus
End Sub
Private Sub RndDices()
    Randomize
    For I = 1 To 6
        If Not Dice(I).Selected Then
            Dice(I).Tag = Int(6 * Rnd + 1)
            Set Dice(I).Die = PictureCnt(Dice(I).Tag).Picture
        End If
    Next
End Sub
Private Function ChkAvailable() As Boolean
Dim D(0 To 2), R(0 To 2), U(0 To 2), M(0 To 2), UD(0 To 2), MX(0 To 2) As Boolean
With Table
    For I = 1 To 3
        D(I - 1) = .SumBox(I).Caption <> vbNullString
        R(I - 1) = .SumBox(20 + I).Caption <> vbNullString
        U(I - 1) = .SumBox(40 + I).Caption <> vbNullString
        M(I - 1) = .SumBox(100 + I).Caption <> vbNullString
        UD(I - 1) = .SumBox(120 + I).Caption <> vbNullString
        MX(I - 1) = .SumBox(140 + I).Caption <> vbNullString
    Next
    If D(0) And D(1) And D(2) And R(0) And R(1) And R(2) And _
       U(0) And U(1) And U(2) And M(0) And M(1) And M(2) And _
       UD(0) And UD(1) And UD(2) And MX(0) And MX(1) And MX(2) Then
        ChkAvailable = False
    Else
        ChkAvailable = True
    End If
End With
End Function
Private Function ChkHand() As Boolean
Dim H(0 To 2) As Boolean
With Table
    For I = 1 To 3
        H(I - 1) = .SumBox(80 + I).Caption <> vbNullString
    Next
    If H(0) And H(1) And H(2) Then
        ChkHand = False
    Else
        ChkHand = True
    End If
End With
End Function
Private Function ChkAnnounce() As Boolean
Dim A(0 To 2) As Boolean
With Table
    For I = 1 To 3
        A(I - 1) = .SumBox(60 + I).Caption <> vbNullString
    Next
    If A(0) And A(1) And A(2) Then
        ChkAnnounce = False
    Else
        ChkAnnounce = True
    End If
End With
End Function
Private Sub AutoClick()
    If Announced Then
        Announce_Click Entered
    Else
        If Not Third Then
            For I = 61 To 73
                If Announce(I).Enabled Then
                    EnableArr.Add I
                    Announce(I).Enabled = False
                End If
            Next
        End If
        Select Case Entered
            Case 0
            Case 1 To 13
                Down_Click Entered
            Case 21 To 33
                Random_Click Entered
            Case 41 To 53
                Up_Click Entered
            Case 101 To 113
                Middle_Click Entered
            Case 121 To 133
                UpDown_Click Entered
            Case 141 To 153
                Max_Click Entered
        End Select
    End If
End Sub
Private Sub Dice_Click(Index As Integer)
    If First And Selected < 5 Then
        If Dice(Index).Selected Then
            Dice(Index).Selected = False
            Selected = Selected - 1
        Else
            Dice(Index).Selected = True
            Selected = Selected + 1
        End If
    Else
        If Dice(Index).Selected Then
            Dice(Index).Selected = False
            Selected = Selected - 1
        End If
    End If
End Sub
Private Sub TurnBtn_Click()
    Select Case Entered
        Case 1 To 12
            Down(Entered).Enabled = False
            Down(Entered + 1).Enabled = True
        Case 13
            Down(13).Enabled = False
        Case 21 To 33
            Random(Entered).Enabled = False
        Case 41
            Up(41).Enabled = False
        Case 42 To 53
            Up(Entered).Enabled = False
            Up(Entered - 1).Enabled = True
        Case 61 To 73
            Announce(Entered).Enabled = False
        Case 81 To 93
            Hand(Entered).Enabled = False
        Case 101
            Middle(101).Enabled = False
        Case 102 To 107
            Middle(Entered).Enabled = False
            Middle(Entered - 1).Enabled = True
        Case 108 To 112
            Middle(Entered).Enabled = False
            Middle(Entered + 1).Enabled = True
        Case 113
            Middle(113).Enabled = False
        Case 121 To 126
            UpDown(Entered).Enabled = False
            UpDown(Entered + 1).Enabled = True
        Case 127 To 128
            UpDown(Entered).Enabled = False
        Case 129 To 133
            UpDown(Entered).Enabled = False
            UpDown(Entered - 1).Enabled = True
        Case 141 To 153
            Max(Entered).Enabled = False
    End Select
    For I = 1 To 6
        Dice(I).Tag = vbNullString
        Set Dice(I).Die = Nothing
        Dice(I).Selected = False
    Next
    For Each Enabler In EnableArr
        If Enabler = Entered Then EnableArr.Remove CStr(Enabler)
    Next
    TurnBtn.Enabled = False
    Announced = False
    mnuUndo.Enabled = False
    Selected = 0
    SelectTransfered = False
    Entered = 0
    First = False
    Second = False
    Third = False
    RollBtn.Enabled = True
    Set RollBtn.Image = DiceCnt(3).Picture
    RollBtn.SetFocus
    Enable
    RefreshStatus
    If Finished Then
        RollBtn.Enabled = False
        NewScore SumBox(0).Caption
        mnuSave.Enabled = False
        mnuBest.Enabled = False
        For I = 1 To 6
            Dice(I).Visible = False
        Next
        SetTimer Me.hwnd, 1, 300, AddressOf HitFlash
        Played = False
        ClearRolls = True
    End If
    RefreshStatus
End Sub
Public Sub LoadBestScores()
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo NoFile
    Set ScoreFile = fso.OpenTextFile(App.Path & "\winYAMB.dat", 1)
    ScoreStr = vbNullString
    Do Until ScoreFile.AtEndOfStream
        ScoreStr = ScoreStr & ScoreFile.ReadLine
    Loop
    ScoreFile.Close
    DecScoreStr = DecryptText(ScoreStr, "r3Y4Na9")
    ScoreArr = Split(DecScoreStr, "#NEW_RECORD#", -1, vbTextCompare)
    For I = 0 To 4
        RecArr = Split(ScoreArr(I), "#RECORD_DATA#", -1, vbTextCompare)
        LabName(I + 1).Caption = RecArr(0)
        LabScore(I + 1).Caption = RecArr(1)
    Next
    Exit Sub
NoFile:
    For I = 1 To 5
        LabName(I).Caption = "---"
        LabScore(I).Caption = "---"
    Next
End Sub
Private Sub NewScore(Score)
    If Len(Score) > 0 Then
        Dim Greater As Boolean
        For I = 1 To 5
            If LabScore(I) <> "---" Then
                If CInt(Score) >= CInt(LabScore(I)) Then Greater = True: NewBest = I: Exit For
            Else
                Greater = True
                NewBest = I
                Exit For
            End If
        Next
        If Greater Then
            If NewBest < 5 Then
                Last = 5
                For I = 1 To 5
                    If LabName(I).Caption = "---" Then Last = I: Exit For
                Next
                For I = Last - 1 To NewBest Step -1
                    LabName(I + 1).Caption = LabName(I).Caption
                    LabScore(I + 1).Caption = LabScore(I).Caption
                Next
            End If
            NewName(NewBest).Visible = True
            NewName(NewBest).Text = vbNullString
            NewName(NewBest).SetFocus
            LabName(NewBest).Caption = vbNullString
            LabScore(NewBest).Caption = Score
            PlayWav "#1"
        Else
            PlayWav "#2"
        End If
    End If
End Sub
Private Sub NewName_Change(Index As Integer)
    LabName(Index).Caption = NewName(Index).Text
End Sub
Private Sub NewName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then NewName(Index).Visible = False: SaveScores
End Sub
Private Sub SaveScores()
    LabName(NewBest).Caption = NewName(NewBest).Text
    SaveStr = vbNullString
    For I = 1 To 5
        SaveStr = SaveStr & LabName(I).Caption & _
                  "#RECORD_DATA#" & LabScore(I).Caption & _
                  "#NEW_RECORD#"
    Next
    SaveStr = Left(SaveStr, Len(SaveStr) - 12)
    EncStr = EncryptText(SaveStr, "r3Y4Na9")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ScoreFile = fso.OpenTextFile(App.Path & "\winYAMB.dat", 2, True)
    ScoreFile.Write EncStr
    ScoreFile.Close
    NewName(NewBest).Visible = False
End Sub
Public Sub StuffMenu()
    hMenu = GetMenu(Me.hwnd)
    SetMenuItemBitmaps hMenu, 2, 0, IcoCnt(1).Picture, 0&
    SetMenuItemBitmaps hMenu, 3, 0, IcoCnt(4).Picture, 0&
    SetMenuItemBitmaps hMenu, 4, 0, IcoCnt(6).Picture, 0&
    SetMenuItemBitmaps hMenu, 6, 0, IcoCnt(7).Picture, 0&
    SetMenuItemBitmaps hMenu, 10, 0, IcoCnt(5).Picture, 0&
    SetMenuItemBitmaps hMenu, 12, 0, IcoCnt(2).Picture, 0&
    SetMenuItemBitmaps hMenu, 14, 1, IcoCnt(3).Picture, 0&
    SetMenuItemBitmaps hMenu, 17, 1, IcoCnt(0).Picture, 0&
End Sub
Public Sub RefreshStatus()
    If mnuUndo.Enabled And Entered <> 0 Then UndoStat.Status = "U" Else UndoStat.Status = vbNullString
    If Not Third And Not ClearRolls Then: If Second Then RollNum = "1" Else: If First Then RollNum = "2" Else RollNum = "3"
    If RollNum Then RollStat.Status = "Rolls left: " & RollNum Else RollStat.Status = vbNullString
    ClearRolls = False
    Empties = 0
    For I = 1 To 153
        Select Case I
            Case 1 To 13
                If Down(I).Text = vbNullString Then Empties = Empties + 1
            Case 21 To 33
                If Random(I).Text = vbNullString Then Empties = Empties + 1
            Case 41 To 53
                If Up(I).Text = vbNullString Then Empties = Empties + 1
            Case 61 To 73
                If Announce(I).Text = vbNullString Then Empties = Empties + 1
            Case 81 To 93
                If Hand(I).Text = vbNullString Then Empties = Empties + 1
            Case 101 To 113
                If Middle(I).Text = vbNullString Then Empties = Empties + 1
            Case 121 To 133
                If UpDown(I).Text = vbNullString Then Empties = Empties + 1
            Case 141 To 153
                If Max(I).Text = vbNullString Then Empties = Empties + 1
        End Select
    Next
    If Empties <> 0 Then TurnStat.Status = "Turns left: " & Empties Else TurnStat.Status = vbNullString
    If First Then
        For Field = 1 To 2
            Value = 0
            Select Case Field
                Case 1
                    A = 1
                    B = 6
                    c = 1
                Case 2
                    A = 6
                    B = 1
                    c = -1
            End Select
            Subtraction = 0
            For I = A To B Step c
                For iter = 1 To 6
                    If CInt(Dice(iter).Tag) = I Then
                        Subtraction = CInt(Dice(iter).Tag)
                        Exit For
                    End If
                Next
                If Subtraction <> 0 Then Exit For
            Next
            For I = 1 To 6
                Value = Value + CInt(Dice(I).Tag)
            Next
            Select Case Field
                Case 1
                    MaxStat.Status = "Max: " & (Value - Subtraction)
                Case 2
                    MinStat.Status = "Min: " & (Value - Subtraction)
            End Select
        Next
    Else
        MaxStat.Status = vbNullString
        MinStat.Status = vbNullString
    End If
End Sub
