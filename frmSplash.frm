VERSION 5.00
Begin VB.Form frmSplash 
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   5775
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picNav 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   0
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   5655
      TabIndex        =   7
      Top             =   600
      Width           =   5655
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   8760
         Y1              =   30
         Y2              =   30
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   108
         X2              =   8760
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox picNav 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   1
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   5655
      TabIndex        =   6
      Top             =   1920
      Width           =   5655
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   108
         X2              =   8760
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   108
         X2              =   8760
         Y1              =   24
         Y2              =   24
      End
   End
   Begin VB.Label lAdding 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   5535
   End
   Begin VB.Label lHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "the"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   5535
   End
   Begin VB.Label lHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "R E X"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   480
      Picture         =   "frmSplash.frx":0000
      Top             =   2400
      Width           =   4530
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Sveinn R. Sigurðsson"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   5535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Copyright (C) 1998 - 2000"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   5535
   End
   Begin VB.Label lVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Label lHeader 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MusicManager"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   5535
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Click()
    Unload Me
End Sub



Private Sub Form_Load()
    lVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision & " Beta"
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub



Private Sub Label1_Click()
    Unload Me
End Sub



Private Sub Label2_Click()
    Unload Me
End Sub



'Private Sub lHeader_Click()
'    Unload Me
'End Sub



Private Sub lVersion_Click()
    Unload Me
End Sub

