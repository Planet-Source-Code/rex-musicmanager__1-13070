Attribute VB_Name = "mPlayList"
Option Explicit



Public PlayListDirectory As String ' Where is our PlayList saved


' Clear the PlayList
Public Sub Clear()
    frmMain.lp.ListItems.Clear
End Sub



' Save Collection
Public Sub SaveFile()
    On Error Resume Next
    If frmMain.lp.ListItems.Count > 0 Then
        Dim i As Long
        frmMain.cd.DialogTitle = "Save Playlist"
        frmMain.cd.ShowSave
        If frmMain.cd.Filename <> "" Then
            Close #1
            If InStr(UCase(frmMain.cd.Filename), ".MUS") > 0 Then
                Kill frmMain.cd.Filename
                Open frmMain.cd.Filename For Output As #1
            Else
                Kill frmMain.cd.Filename
                Open frmMain.cd.Filename & ".MUS" For Output As #1
            End If
                Write #1, frmMain.lp.ListItems.Count
            For i = 1 To frmMain.lp.ListItems.Count
                Write #1, frmMain.lp.ListItems(i).Text ' PlayList
                Write #1, frmMain.lp.ListItems(i).ListSubItems(1).Text ' Rate
                Write #1, frmMain.lp.ListItems(i).ListSubItems(2).Text  ' FileName
                Write #1, frmMain.lp.ListItems(i).ListSubItems(3).Text  ' Path
            Next i
        End If
    Else
        Dim response As String
        response = MsgBox("There are no songs on the Playlist", vbCritical + vbOKOnly, "Error")
    End If
End Sub



' Open PlayList collection
Public Sub OpenPlayList()
On Error Resume Next
    frmMain.cd.DialogTitle = "Open Playlist"
    frmMain.cd.ShowOpen
    If frmMain.cd.Filename <> "" Then
        If frmMain.lp.ListItems.Count > 0 Then
            Dim response As String
            response = MsgBox("Would you like to add the contents of " & mEngine.ExtractFileName(frmMain.cd.Filename) & " to the existing Playlist?", vbQuestion + vbYesNo, "PlayList")
            If response = 7 Then frmMain.lp.ListItems.Clear
        End If
        Call ReadMUSFormat(frmMain.cd.Filename)
        Call mInterface.ShowPlayList
    End If
End Sub



' Read a MUS format file
Private Sub ReadMUSFormat(sFilenameMUS As String)
        Dim i As Long
        Close #1
        Dim t As String
        Dim iCount As Long
        Dim iCount2 As Long
        iCount2 = 1
        Open sFilenameMUS For Input As #1
        Input #1, iCount
        DoEvents
            For i = 1 To iCount
                Input #1, t
                If t = "" Then t = " "
                frmMain.lp.ListItems.add = t
                ' Rate
                Input #1, t
                If t = "" Then t = " "
                frmMain.lp.ListItems(frmMain.lp.ListItems.Count).ListSubItems.add , , t
                ' FileName
                Input #1, t
                If t = "" Then t = " "
                frmMain.lp.ListItems(frmMain.lp.ListItems.Count).ListSubItems.add , , t
                ' Path
                Input #1, t
                If t = "" Then t = " "
                frmMain.lp.ListItems(frmMain.lp.ListItems.Count).ListSubItems.add , , t
                iCount2 = iCount2 + 1
            Next i
        Close #1
    frmMain.sbrMain.Panels(2).Text = frmMain.lp.ListItems.Count & " songs"
End Sub


