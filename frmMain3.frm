VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "MusicManager"
   ClientHeight    =   7590
   ClientLeft      =   165
   ClientTop       =   405
   ClientWidth     =   9660
   Icon            =   "frmMain3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lp 
      Height          =   4575
      Left            =   720
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "P"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "R"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Filename"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Path"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Artist"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Album"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Genre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Length"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "size"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lf 
      Height          =   4575
      Left            =   360
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "P"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "R"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Filename"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Path"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Artist"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Album"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Genre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Length"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "size"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame fFind 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   9615
      Begin MSComctlLib.Toolbar cFind 
         Height          =   360
         Left            =   3120
         TabIndex        =   17
         Top             =   1200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   635
         ButtonWidth     =   1429
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Find"
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.TextBox tAlbum 
         Height          =   285
         Left            =   5400
         TabIndex        =   3
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox tYear 
         Height          =   285
         Left            =   5400
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cGenre 
         Height          =   315
         Left            =   8160
         TabIndex        =   4
         Text            =   "Genre..."
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox tTitle 
         Height          =   285
         Left            =   600
         TabIndex        =   0
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox tPath 
         Height          =   285
         Left            =   600
         TabIndex        =   1
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label lGenre 
         Caption         =   "Genre"
         Height          =   255
         Left            =   7560
         TabIndex        =   16
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lAlbum 
         Caption         =   "Album"
         Height          =   255
         Left            =   4800
         TabIndex        =   15
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lYear 
         Caption         =   "Year"
         Height          =   255
         Left            =   4800
         TabIndex        =   14
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lFindHeader 
         BackColor       =   &H8000000C&
         Caption         =   " Music Collection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   9495
      End
      Begin VB.Label lTitle 
         Caption         =   "Title"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lPath 
         Caption         =   "Path "
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
   End
   Begin MSComctlLib.ImageList iL 
      Left            =   8880
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain3.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain3.frx":1396
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain3.frx":3B4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain3.frx":4616
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain3.frx":52F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain3.frx":5FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain3.frx":6CA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   7335
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13970
            Text            =   "Total files"
            TextSave        =   "Total files"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView l 
      Height          =   4575
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "P"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "R"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Filename"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Path"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Artist"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Album"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Genre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Length"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "size"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   1429
      ButtonWidth     =   1905
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "iL"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Collection"
            Key             =   "AddCD"
            Object.ToolTipText     =   "Add CD to database"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "aCD"
                  Text            =   "Add &CD"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "aDrive"
                  Text            =   "Add &Drive"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "aDir"
                  Text            =   "Add D&irectory"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "aFile"
                  Text            =   "Add &file"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&PlayList"
            Key             =   "AddDir"
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NewPlay"
                  Text            =   "&New"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "OpenPlay"
                  Text            =   "&Open"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PlaySep1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PlaySave"
                  Text            =   "&Save"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Find"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Favourites"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Preferences"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MediaPlayerCtl.MediaPlayer mp 
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Line topLine 
      BorderColor     =   &H00808080&
      X1              =   -240
      X2              =   9600
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
      End
   End
   Begin VB.Menu mnuPlayList 
      Caption         =   "&PlayList"
      Begin VB.Menu mnuPrevious 
         Caption         =   "&Previous"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuNext 
         Caption         =   "&Next Song"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "&Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add to Playlist"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove selection"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




' Find Data in Collection
Private Sub cFind_ButtonClick(ByVal Button As MSComctlLib.Button)
    Call mFind.Find
End Sub



' Forms Startup routine
Private Sub Form_Load()
    If mStartup.sCommandLine <> "" Then mStartup.CommandLine
    mMusicDevices.iCol = 1
End Sub



' Resize the MainForm
Private Sub Form_Resize()
    Call mInterface.resize(Me)
End Sub



' Play the selected song from the Collection tab
Private Sub l_DblClick()
On Error GoTo e:
    mMusicDevices.iCol = 1
    mMusicDevices.t = l.SelectedItem.Index
    Call mMusicDevices.PlaySong(Me)
e:
    Exit Sub
End Sub



' Display the Popupmenu
Private Sub l_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub



' Play the selected song from the find tab
Private Sub lf_DblClick()
On Error GoTo e:
    mMusicDevices.iCol = 3
    mMusicDevices.t = lf.SelectedItem.Index
    Call mMusicDevices.PlaySong(Me)
e:
    Exit Sub
End Sub



' Show Popupmenu
Private Sub lf_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub



' Play from the PlayList collection
Private Sub lp_DblClick()
On Error GoTo e:
    mMusicDevices.iCol = 2
    mMusicDevices.t = lp.SelectedItem.Index
    Call mMusicDevices.PlaySong(Me)
e:
    Exit Sub
End Sub



' Handle KeyPress for PlayList collection
Private Sub lp_KeyPress(keyascii As Integer)
    Call mInterface.HandleKeyPress(keyascii)
End Sub



' Display a popupmenu
Private Sub lp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub



' Display the Splash Window
Private Sub mnuAbout_Click()
    mInterface.ShowSplash
End Sub



' Add To PlayList
Private Sub mnuAdd_Click()
    Call mInterface.AddToPlayList
End Sub



' Quit MusicManager
Private Sub mnuExit_Click()
    Call mInterface.ExitMM
End Sub



' Create new MusicCollection
Private Sub mnuNew_Click()
    frmNewCollection.Show vbModal
End Sub



' Load the next song and play it
Private Sub mnuNext_Click()
    Call mMusicDevices.nextSong(Me)
    Call mMusicDevices.Play(Me)
End Sub




' Open DataBase from Disk
Private Sub mnuOpen_Click()
    Call mInterface.OpenFile(Me)
End Sub



' Play the selected song
Private Sub mnuPlay_Click()
    Call mMusicDevices.Play(Me)
End Sub




' Play the previous song in a collection
Private Sub mnuPrevious_Click()
    Call mMusicDevices.previousSong(Me)
End Sub



' Remove Selection from the selected Collection
Private Sub mnuRemove_Click()
    Call mInterface.RemoveSong
End Sub



' Save Collection to disk
Private Sub mnuSave_Click()
    Call mInterface.SaveFile(Me)
End Sub



' Stop Playing the selected song
Private Sub mnuStop_Click()
    Call mMusicDevices.StopDevice(Me)
End Sub



' Play Next Song on PlayList
Private Sub mp_EndOfStream(ByVal Result As Long)
    On Error GoTo e:
        mMusicDevices.SelectSong (False)
        mMusicDevices.t = mMusicDevices.t + 1
        Call mMusicDevices.PlaySong(Me)
        mMusicDevices.SelectSong (True)
        Exit Sub
e:
        mMusicDevices.t = 1
        Call mMusicDevices.PlaySong(Me)
        Call mMusicDevices.SelectSong(True)
End Sub



' User pressed the toolbar
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    ' User pressed the Collection button
    If Button.Index = 1 Then
        Call ShowCollection
    ElseIf Button.Index = 3 Then
        Call mInterface.ShowPlayList
    ElseIf Button.Index = 4 Then
        l.Visible = False
        lp.Visible = False
        lf.Visible = True
        lFindHeader.Caption = " Find Music"
        mMusicDevices.iCol = 3
    End If
End Sub


' User Pressed the Collection's SubMenu
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Key = "aDrive" Then mCollection.AddDir
    If ButtonMenu.Key = "NewPlay" Then mPlayList.Clear
    If ButtonMenu.Key = "PlaySave" Then mPlayList.SaveFile
    If ButtonMenu.Key = "OpenPlay" Then mPlayList.OpenPlayList
End Sub



' User typed in a path
Private Sub tPath_Change()
    If tPath.Text <> "" Then cFind.Enabled = True
End Sub



' User pressed a key in tPath
Private Sub tPath_KeyPress(keyascii As Integer)
    If keyascii = 13 And tPath.Text <> "" Then Call mFind.Find
End Sub



' User typed in a title
Private Sub tTitle_Change()
    If tTitle.Text <> "" Then cFind.Enabled = True
End Sub



' User Pressed the Key in tTitle
Private Sub tTitle_KeyPress(keyascii As Integer)
    If keyascii = 13 And tTitle.Text <> "" Then Call mFind.Find
End Sub
