Attribute VB_Name = "mMusicDevices"
Option Explicit

Public t            As Long         ' Current song index in the Collection
Public iCol         As Integer      ' Which Collection are we playing from



' Play next song in collection or playlist
Public Sub nextSong(ff As Form)
    On Error GoTo e:
        Call SelectSong(False)
        t = t + 1
        Call SelectSong(True)
        Call PlaySong
        Exit Sub
e:
        t = 1
        Call PlaySong
End Sub



' Play Previous song in a collection or playlist
Public Sub previousSong(ff As Form)
On Error GoTo e:
    Call SelectSong(False)
    t = t - 1
    Call SelectSong(True)
    Call PlaySong
    Exit Sub
e:
    If iCol = 1 Then t = frmMain.l.ListItems.Count
    If iCol = 3 Then t = frmMain.lf.ListItems.Count
End Sub



' Play song by index t, from a collection iCol
Public Sub PlaySong()
On Error GoTo e:
    ' Find out what list to play
    If iCol = 1 Then
        If Len(frmMain.l.ListItems(t).ListSubItems(3).Text) > 3 Then
            frmMain.mp.Filename = frmMain.l.ListItems(t).ListSubItems(3).Text & "\" & frmMain.l.ListItems(t).ListSubItems(2).Text
        Else
            frmMain.mp.Filename = frmMain.l.ListItems(t).ListSubItems(3).Text & frmMain.l.ListItems(t).ListSubItems(2).Text
        End If
        frmMain.sbrMain.Panels(1).Text = "Now playing : " & mMusicEngine.ExtractFileName(frmMain.mp.Filename)
        DoEvents
    ElseIf iCol = 2 Then
        If Len(frmMain.lp.ListItems(t).ListSubItems(3).Text) > 3 Then
            frmMain.mp.Filename = frmMain.lp.ListItems(t).ListSubItems(3).Text & "\" & frmMain.lp.ListItems(t).ListSubItems(2).Text
        Else
            frmMain.mp.Filename = frmMain.lp.ListItems(t).ListSubItems(3).Text & frmMain.lp.ListItems(t).ListSubItems(2).Text
        End If
        frmMain.sbrMain.Panels(1).Text = "Now playing : " & mMusicEngine.ExtractFileName(frmMain.mp.Filename)
    ElseIf iCol = 3 Then
        If Len(frmMain.lf.ListItems(t).ListSubItems(3).Text) > 3 Then
            frmMain.mp.Filename = frmMain.lf.ListItems(t).ListSubItems(3).Text & "\" & frmMain.lf.ListItems(t).ListSubItems(2).Text
        Else
            frmMain.mp.Filename = frmMain.lf.ListItems(t).ListSubItems(3).Text & frmMain.lf.ListItems(t).ListSubItems(2).Text
        End If
        frmMain.sbrMain.Panels(1).Text = "Now playing : " & mMusicEngine.ExtractFileName(frmMain.mp.Filename)
    End If
e:
    Exit Sub
End Sub



' Play Song
Public Sub Play(ff As Form)
On Error GoTo e:
    ff.mp.Stop
    ff.mp.Play
e:
    Exit Sub
End Sub



' Stop Song
Public Sub StopDevice(ff As Form)
    ff.mp.Stop
End Sub



' Highlight the currently playing song
Public Sub SelectSong(b As Boolean)
    If iCol = 1 Then frmMain.l.ListItems.Item(t).Selected = b
    If iCol = 2 Then frmMain.lp.ListItems.Item(t).Selected = b
    If iCol = 3 Then frmMain.lf.ListItems.Item(t).Selected = b
End Sub

