Attribute VB_Name = "mEngine"
Option Explicit



' Extract filename from path + filename
Public Function ExtractFileName(sFileName As String) As String
    Dim l As Long
    Dim i As Long
    l = Len(sFileName)
    
    For i = 1 To l
        If InStr(Right(sFileName, i), "\") > 0 Then
            sFileName = Right(sFileName, i - 1)
            Exit For
        End If
    Next i
    ExtractFileName = sFileName
End Function
