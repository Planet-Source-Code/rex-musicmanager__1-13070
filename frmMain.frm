VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Svenni's MP3 collection"
   ClientHeight    =   6585
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8745
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   1005
      ButtonWidth     =   2117
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add &CD"
            Key             =   "AddCD"
            Object.ToolTipText     =   "Add CD to database"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add &Directory"
            Key             =   "AddDir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add &file"
            Key             =   "AddFile"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
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
      Begin VB.Menu mnuFirst 
         Caption         =   "&First song"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuPrev 
         Caption         =   "P&revious song"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuNext 
         Caption         =   "&Next song"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuLast 
         Caption         =   "&Last song"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlay 
         Caption         =   "&Play"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "&Stop"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuAddToPlayList 
         Caption         =   "&Add to PlayList"
      End
      Begin VB.Menu mnuRate 
         Caption         =   "&Rate"
         Begin VB.Menu mnu100 
            Caption         =   "100%"
         End
         Begin VB.Menu mnu95 
            Caption         =   "95%"
         End
         Begin VB.Menu mnu90 
            Caption         =   "90%"
         End
         Begin VB.Menu mnu85 
            Caption         =   "85%"
         End
         Begin VB.Menu mnu80 
            Caption         =   "80%"
         End
         Begin VB.Menu mnu75 
            Caption         =   "75%"
         End
         Begin VB.Menu mnu70 
            Caption         =   "70%"
         End
         Begin VB.Menu mnu65 
            Caption         =   "65%"
         End
         Begin VB.Menu mnu60 
            Caption         =   "60%"
         End
         Begin VB.Menu mnu55 
            Caption         =   "55%"
         End
         Begin VB.Menu mnu50 
            Caption         =   "50%"
         End
         Begin VB.Menu mnu40 
            Caption         =   "40%"
         End
         Begin VB.Menu mnu35 
            Caption         =   "35%"
         End
         Begin VB.Menu mnu30 
            Caption         =   "30%"
         End
         Begin VB.Menu mnu25 
            Caption         =   "25%"
         End
         Begin VB.Menu mnu20 
            Caption         =   "20%"
         End
         Begin VB.Menu mnu15 
            Caption         =   "15%"
         End
         Begin VB.Menu mnu10 
            Caption         =   "10%"
         End
         Begin VB.Menu mnu5 
            Caption         =   "5%"
         End
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "R&emove"
      End
      Begin VB.Menu mnuSongProp 
         Caption         =   "&Song Properties"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    frmDocument.Show
End Sub

Private Sub MDIForm_Terminate()
    frmDocument.mp.Stop
    End
End Sub

Private Sub mnu10_Click()
    Call rate(10)
End Sub

Private Sub mnu100_Click()
    Call rate(100)
End Sub

Private Sub mnu15_Click()
    Call rate(15)
End Sub

Private Sub mnu20_Click()
    Call rate(20)
End Sub

Private Sub mnu25_Click()
    Call rate(25)
End Sub

Private Sub mnu30_Click()
    Call rate(30)
End Sub

Private Sub mnu35_Click()
    Call rate(35)
End Sub

Private Sub mnu40_Click()
    Call rate(40)
End Sub

Private Sub mnu5_Click()
    Call rate(5)
End Sub

Private Sub mnu50_Click()
    Call rate(50)
End Sub

Private Sub mnu55_Click()
    Call rate(55)
End Sub

Private Sub mnu60_Click()
    Call rate(60)
End Sub

Private Sub mnu65_Click()
    Call rate(65)
End Sub

Private Sub mnu70_Click()
    Call rate(70)
End Sub

Private Sub mnu75_Click()
    Call rate(75)
End Sub

Private Sub mnu80_Click()
    Call rate(80)
End Sub

Private Sub mnu85_Click()
    Call rate(85)
End Sub

Private Sub mnu90_Click()
    Call rate(90)
End Sub

Private Sub mnu95_Click()
    Call rate(95)
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuFind_Click()
    frmFind.Show vbModal
End Sub

Private Sub mnuFirst_Click()
    Let frmDocument.t = -1
    Call frmDocument.nextSong
End Sub

Private Sub mnuLast_Click()
    frmDocument.t = frmDocument.l.ListItems.Count - 1
    Call frmDocument.nextSong
End Sub

Private Sub mnuNext_Click()
    Call frmDocument.nextSong
End Sub



Private Sub mnuPlay_Click()
On Error Resume Next
    frmDocument.mp.Play
End Sub

Private Sub mnuPrev_Click()
    frmDocument.t = frmDocument.t - 2
    Call frmDocument.nextSong
End Sub

Private Sub mnuSave_Click()
    cd.DialogTitle = "Save database"
    cd.ShowSave
    If cd.FileName <> "" Then
        Close #1
        If InStr(UCase(cd.FileName), ".MUS") > 0 Then
            Open cd.FileName For Output As #1
        Else
            Open cd.FileName & ".MUS" For Output As #1
        End If
            Write #1, frmDocument.l.ListItems.Count
        For i = 1 To frmDocument.l.ListItems.Count
            Write #1, frmDocument.l.ListItems(i).Text ' PlayList
            Write #1, frmDocument.l.ListItems(i).ListSubItems(1).Text ' Rate
            Write #1, frmDocument.l.ListItems(i).ListSubItems(2).Text  ' FileName
            Write #1, frmDocument.l.ListItems(i).ListSubItems(3).Text  ' Path
        Next i
    End If
End Sub

Private Sub mnuStop_Click()
    frmDocument.mp.Stop
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 2 Then Call addDirectory
End Sub

Sub addDirectory()
    frmAddDir.Show vbModal
    Call AddFiles
End Sub


Public Sub AddFiles()
    For i = 0 To frmAddDir.File1.ListCount - 1
        frmDocument.l.ListItems.Add = ""
        frmDocument.l.ListItems(frmDocument.l.ListItems.Count).ListSubItems.Add , , ""
        frmDocument.l.ListItems(frmDocument.l.ListItems.Count).ListSubItems.Add , , frmAddDir.f.List(i) '"Best í bílinn - #2"
        frmDocument.l.ListItems(frmDocument.l.ListItems.Count).ListSubItems.Add , , frmAddDir.Dir1.Path
        frmDocument.l.ListItems(frmDocument.l.ListItems.Count).ListSubItems.Add , , "Sveinn R. Sigurðsson"
        frmDocument.l.ListItems(frmDocument.l.ListItems.Count).ListSubItems.Add , , "Perfect Play of Life"
        frmDocument.l.ListItems(frmDocument.l.ListItems.Count).ListSubItems.Add , , "Upgrade"
        frmDocument.l.ListItems(frmDocument.l.ListItems.Count).ListSubItems.Add , , "2000"
        frmDocument.l.ListItems(frmDocument.l.ListItems.Count).ListSubItems.Add , , "4:10"
        frmDocument.l.ListItems(frmDocument.l.ListItems.Count).ListSubItems.Add , , "D:\Best of pop\DUO\ppl.mp3"
        frmDocument.l.ListItems(frmDocument.l.ListItems.Count).ListSubItems.Add , , "4.103 kb"
    Next i
End Sub

Public Sub rate(iPercentage As Integer)
On Error Resume Next
    With frmDocument.l
        .ListItems(.SelectedItem.Index).ListSubItems(1).Text = iPercentage & "%"
    End With
End Sub
