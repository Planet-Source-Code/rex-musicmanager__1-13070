Attribute VB_Name = "mInterface"
Option Explicit



Public Sub resize(ff As Form)
On Error Resume Next
    With frmMain
        ' Resize the listViews
        .l.Move .l.left, .topLine.Y1 + 80, .ScaleWidth - (.l.left * 2), .ScaleHeight - .topLine.Y1 - .sbrMain.Height - 100
        .lf.Move .l.left, .topLine.Y1 + 80, .ScaleWidth - (.l.left * 2), .ScaleHeight - .topLine.Y1 - .sbrMain.Height - 100
        .lp.Move .l.left, .topLine.Y1 + 80, .ScaleWidth - (.l.left * 2), .ScaleHeight - .topLine.Y1 - .sbrMain.Height - 100
        .lt.Move .l.left, .topLine.Y1 + 80, .ScaleWidth - (.l.left * 2), .ScaleHeight - .topLine.Y1 - .sbrMain.Height - 100
        ' Resize Find Panels
        .fFind.Move 0, .fFind.tOp, .ScaleWidth, .fFind.Height
        .lFindHeader.Move .lTitle.left, .lFindHeader.tOp, .ScaleWidth - (.lTitle.left * 2), .lFindHeader.Height
        .cGenre.left = .ScaleWidth - .cGenre.Width - .lFindHeader.left
        .lGenre.left = .ScaleWidth - .cGenre.Width - .lFindHeader.left - .lGenre.Width
        .tAlbum.left = .lGenre.left - .tAlbum.Width - 400
        .lAlbum.left = .tAlbum.left - .lAlbum.Width
        .tYear.left = .tAlbum.left
        .lYear.left = .tAlbum.left - .lYear.Width
        .tTitle.Width = .lYear.left - .tTitle.left - 400
        .tPath.Width = .lYear.left - .tPath.left - 400
        .cFind.left = .tPath.Width + .tPath.left - .cFind.Width
        ' Set TopLine Width
        .topLine.X1 = .l.left
        .topLine.X2 = .ScaleWidth - (.l.left * 2)
        .Player.Width = .ScaleWidth - (.Player.left * 2)
    End With
End Sub



' Open DataBase
Public Sub OpenFile(ff As Form)
On Error Resume Next
    frmMain.cd.DialogTitle = "Open music database"
    frmMain.cd.ShowOpen
    If frmMain.cd.Filename <> "" And InStr(UCase(frmMain.cd.Filename), ".MUS") > 0 Then
        Call ReadMUSFormat(frmMain.cd.Filename)
    ElseIf frmMain.cd.Filename <> "" And InStr(UCase(frmMain.cd.Filename), ".REX") > 0 Then
        Call mDataBase.ReadREXFormat(frmMain.cd.Filename)
    Else
        Dim response As String
        response = MsgBox("Not a valid MusicManager file", vbCritical + vbOKOnly, "Error")
    End If
End Sub



' Read a MUS format file
Public Sub ReadMUSFormat(sFilenameMUS As String)
        Dim i As Long
        Close #1
        Dim t As String
        Dim iCount As Long
        Dim iCount2 As Long
        iCount2 = 1
        Open sFilenameMUS For Input As #1
        Call ClearCollections
        Input #1, iCount
        DoEvents
            For i = 1 To iCount
                Input #1, t
                If t = "" Then t = " "
                ' Preview available
                frmMain.l.ListItems.add = t
                ' Rate
                Input #1, t
                If t = "" Then t = " "
                frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , t
                ' FileName
                Input #1, t
                If t = "" Then t = " "
                frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , t
                ' Path
                Input #1, t
                If t = "" Then t = " "
                frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , t
                If frmSplash.Visible = True Then Call ShowAddingInSplash(t)
                ' Artist
                Input #1, t
                If t = "" Then t = " "
                frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , t
                ' Title
                Input #1, t
                If t = "" Then t = " "
                frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , t
                ' Album
                Input #1, t
                If t = "" Then t = " "
                frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , t
                ' Year
                Input #1, t
                If t = "" Then t = " "
                frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , t
                ' Genre
                Input #1, t
                If t = "" Then t = " "
                frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , t
                iCount2 = iCount2 + 1
            Next i
        Close #1
        frmMain.sbrMain.Panels(2).Text = frmMain.l.ListItems.Count & " songs"
End Sub



' Save Collection
Public Sub SaveFile(ff As Form)
On Error Resume Next
    Dim i As Long
    ff.cd.DialogTitle = "Save database"
    ff.cd.ShowSave
    DoEvents
    If ff.cd.Filename <> "" Then
        Kill ff.cd.Filename
        Close #1
        If InStr(UCase(ff.cd.Filename), ".MUS") > 0 Then
            Open ff.cd.Filename For Output As #1
        Else
            Open ff.cd.Filename & ".MUS" For Output As #1
        End If
            Write #1, ff.l.ListItems.Count
        For i = 1 To ff.l.ListItems.Count
            Write #1, ff.l.ListItems(i).Text                  ' PlayList
            Write #1, ff.l.ListItems(i).ListSubItems(1).Text  ' Rate
            Write #1, ff.l.ListItems(i).ListSubItems(2).Text  ' FileName
            Write #1, ff.l.ListItems(i).ListSubItems(3).Text  ' Path
            Write #1, ff.l.ListItems(i).ListSubItems(4).Text  ' Artist
            Write #1, ff.l.ListItems(i).ListSubItems(5).Text  ' Title
            Write #1, ff.l.ListItems(i).ListSubItems(6).Text  ' Album
            Write #1, ff.l.ListItems(i).ListSubItems(7).Text  ' Year
            Write #1, ff.l.ListItems(i).ListSubItems(8).Text  ' Genre
        Next i
    End If
End Sub



' Exit MusicManager
Public Sub ExitMM()
    End
End Sub



' Show the Collection Panel
Public Sub ShowCollection()
    With frmMain
        .l.Visible = True
        .lp.Visible = False
        .lf.Visible = False
        .lt.Visible = False
        mMusicDevices.iCol = 1
        .lFindHeader.Caption = " Music Collection"
        .sbrMain.Panels(2).Text = .l.ListItems.Count - 1 & " files"
    End With
End Sub



' Clear All Collections
Public Sub ClearCollections()
    frmMain.l.ListItems.Clear
    frmMain.lf.ListItems.Clear
End Sub



' Display the Splash Window
Public Sub ShowSplash()
    frmSplash.lVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    frmSplash.Show vbModal
    Set frmSplash = Nothing
End Sub



' Show the PlayList Collection
Public Sub ShowPlayList()
    frmMain.l.Visible = False
    frmMain.lf.Visible = False
    frmMain.lp.Visible = True
    frmMain.lt.Visible = False
    frmMain.lFindHeader.Caption = " Playlist"
    frmMain.sbrMain.Panels(2).Text = frmMain.lp.ListItems.Count - 1 & " files on Playlist"
End Sub



' Handle keypress for all collection
Public Sub HandleKeyPress(keyascii As Integer)
    If keyascii = 23 Then
        Call RemoveSong
    End If
End Sub



' Add selected data to the PlayList collection
Public Sub AddToPlayList()
    Dim oLV As ListView
    Dim i As Long
    
    frmMain.MousePointer = 13
    With frmMain
        If iCol = 3 Then
        For i = 1 To .lf.ListItems.Count
        
        
            If .lf.ListItems.Item(i).Selected = True Then
                    .lp.ListItems().add = " "
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , ""
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .lf.ListItems(i).ListSubItems(2).Text
                    If .lt.Nodes.Item("LPADDMP3").Checked = True Then
                        .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , App.Path & "\My Music"
                    Else
                        .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .lf.ListItems(i).ListSubItems(3).Text
                    End If
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .lf.ListItems(i).ListSubItems(4).Text
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .lf.ListItems(i).ListSubItems(5).Text
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .lf.ListItems(i).ListSubItems(6).Text
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .lf.ListItems(i).ListSubItems(7).Text
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .lf.ListItems(i).ListSubItems(8).Text
                    '.lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .lf.ListItems(i).ListSubItems(9).Text
                    '.lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .lf.ListItems(i).ListSubItems(10).Text
                    Call mShell.CopyFile(i, mPlayList.PlayListDirectory)
                End If
            Next i
        ElseIf iCol = 1 Then
        For i = 1 To .l.ListItems.Count
            If .l.ListItems.Item(i).Selected = True Then
                    .lp.ListItems().add = " "
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , ""
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .l.ListItems(i).ListSubItems(2).Text
                    If .lt.Nodes.Item("LPADDMP3").Checked = True Then
                        .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , App.Path & "\Playlist"
                    Else
                        .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .l.ListItems(i).ListSubItems(3).Text
                    End If
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .l.ListItems(i).ListSubItems(4).Text
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .l.ListItems(i).ListSubItems(5).Text
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .l.ListItems(i).ListSubItems(6).Text
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .l.ListItems(i).ListSubItems(7).Text
                    .lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , .l.ListItems(i).ListSubItems(8).Text
                    '.lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , "D:\Best of pop\DUO\ppl.mp3"
                    '.lp.ListItems(.lp.ListItems.Count).ListSubItems.add , , "4.103 kb"
                    Call mShell.CopyFile(i, mPlayList.PlayListDirectory)
                End If
            Next i
        End If
    End With
    frmMain.MousePointer = 0
End Sub



' Remove selected songs from the selected collection
Public Sub RemoveSong()
On Error Resume Next ' Error occurs when we can't delete the original file
    With frmMain
        Dim i As Long
        Dim tempKillName As String
        For i = .lp.ListItems.Count To 1 Step -1
            If .lp.ListItems.Item(i).Selected = True Then
                If .lt.Nodes.Item("LPDELFILE").Checked = True Then
                    tempKillName = .lp.ListItems(i).ListSubItems(3).Text & "\" & .lp.ListItems(i).ListSubItems(2).Text
                    Kill tempKillName
                End If
                .lp.ListItems.Remove (i)
            End If
        Next i
    End With
End Sub



' Get Volume information from a drive
Public Sub GetDriveVolume(sVolume As String)

   Dim r As Long
   Dim PathName As String
   Dim DrvVolumeName As String
   Dim DrvSerialNo As String

  'the drive to check
   PathName$ = "d:\"
  
   rgbGetVolume PathName, DrvVolumeName, DrvSerialNo

  'show the results
   
   'frmMain.Caption = "  Drive Statistics for  :  " & UCase$(PathName)
   'frmMain.Caption = "  Volume Label " & DrvVolumeName
   'frmMain.Caption = "  Volume Serial No " & DrvSerialNo

End Sub


Private Sub rgbGetVolume(PathName As String, _
                         DrvVolumeName As String, _
                         DrvSerialNo As String)
 
  'create working variables
  'to keep it simple, use dummy variables for info
  'we're not interested in right now
   Dim r As Long
   Dim pos As Integer
   Dim hword As Long
   Dim HiHexStr As String
   Dim lword As Long
   Dim LoHexStr As String
   Dim VolumeSN As Long
   Dim MaxFNLen As Long

   Dim UnusedStr As String
   Dim UnusedVal1 As Long
   Dim UnusedVal2 As Long

  'pad the strings
   DrvVolumeName$ = Space$(14)
   UnusedStr$ = Space$(32)

  'do what it says
   r = GetVolumeInformation(PathName, _
                            DrvVolumeName, _
                            Len(DrvVolumeName), _
                            VolumeSN&, _
                            UnusedVal1, UnusedVal2, _
                            UnusedStr, Len(UnusedStr$))


  'error check
   If r& = 0 Then Exit Sub

  'determine the volume label
   pos = InStr(DrvVolumeName, Chr$(0))
   If pos Then DrvVolumeName = left$(DrvVolumeName, pos - 1)
   If Len(Trim$(DrvVolumeName)) = 0 Then DrvVolumeName = "(no label)"

  'determine the drive volume id
   hword = HiWord(VolumeSN)
   lword = LoWord(VolumeSN)
   HiHexStr = Format$(Hex(hword), "0000")
   LoHexStr = Format$(Hex(lword), "0000")
 
   DrvSerialNo = HiHexStr & "-" & LoHexStr

End Sub



Private Function HiWord(dw As Long) As Integer
On Error Resume Next
    HiWord = (dw And &HFFFF0000) \ &H10000
End Function
  
  

Private Function LoWord(dw As Long) As Integer
On Error Resume Next
    If dw And &H8000& Then
        LoWord = dw Or &HFFFF0000
    Else
        LoWord = dw And &HFFFF&
    End If
End Function



' If there is a commandline then show the add process in the splash window
Private Sub ShowAddingInSplash(t As String)
    If frmSplash.Visible = True Then
        If t <> "" Or t <> " " Then frmSplash.lAdding.Caption = "Adding : " & t & " ..."
        DoEvents
    End If
End Sub



' Show the Preference tree
Public Sub ShowPreference()
    With frmMain
        .l.Visible = False
        .lf.Visible = False
        .lp.Visible = False
        .lt.Visible = True
        .lFindHeader.Caption = " Preferences"
    End With
End Sub



' Initialize MusicManager on Startup
Public Sub Initialize()
    If mStartup.sCommandLine <> "" Then mStartup.CommandLine
    mMusicDevices.iCol = 1
    mPreferences.CreatePreferences
    mPlayList.PlayListDirectory = App.Path & "\PlayList"
    mPreferences.ChangeBackGroundColor (frmMain.tColor.BackColor)
End Sub
