Attribute VB_Name = "mMusicEngine"
Option Explicit



' Return a filename without a path
Public Function extractFilename(sFile As String) As String
    Dim l As Long
    Dim i As Long
    l = Len(sFile)
    For i = 1 To l
        If InStr(Right(sFile, i), "\") > 0 Then
            sFile = Right(sFile, i - 1)
            Exit For
        End If
    Next i
    extractFilename = sFile
End Function



' Return a path without the filename
Public Function extractPath(sFile As String) As String
    Dim l As Long
    Dim i As Long
    l = Len(sFile)
    For i = 1 To l
        If InStr(Right(sFile, i), "\") > 0 Then
            sFile = Left(sFile, (Len(sFile)) - i)
            Exit For
        End If
    Next i
    extractPath = sFile
End Function

