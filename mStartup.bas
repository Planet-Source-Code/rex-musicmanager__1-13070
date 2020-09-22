Attribute VB_Name = "mStartup"
Option Explicit



Public sCommandLine As String



' Programs startup routine
Sub Main()
On Error GoTo e:

    Dim PROJECTID As String
    PROJECTID = Command()
    If Command() <> "" Then
        sCommandLine = PROJECTID
    End If
    frmMain.Visible = True
e:
  Exit Sub
End Sub



' Handle the commandline parameter
' Is there anything in the commandline parameter
Public Sub CommandLine()
On Error Resume Next
    Dim lengd       As Long
    Dim CommandLine As String
    Dim oldCap      As String
    Dim sFile       As String
    CommandLine = mStartup.sCommandLine
    lengd = Len(CommandLine)
    If CommandLine <> "" Then
        If InStr(UCase(CommandLine), ".MUS") Then
            frmSplash.Show
            Call mInterface.ReadMUSFormat(sCommandLine)
            Unload frmSplash
            Set frmSplash = Nothing
            mMusicDevices.iCol = 1
            mMusicDevices.t = 1
            Call mMusicDevices.PlaySong
        End If
    End If
e:
    Exit Sub
End Sub




