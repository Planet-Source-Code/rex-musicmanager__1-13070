Attribute VB_Name = "mDataBase"
Option Explicit


' Add Current MusicCollection to a database
Public Sub AddCollection()
On Error Resume Next
 Dim i As Long
 Dim db As Database
 Dim dbrecordset As Recordset
 Dim td As TableDef
 Set db = DBEngine.Workspaces(0).OpenDatabase(mVariables.musicDataBase)
 Set dbrecordset = db.OpenRecordset("MusicCollection")
 
 If dbrecordset.RecordCount > 0 Then
  dbrecordset.MoveLast
 End If
 
 With dbrecordset
 For i = 1 To frmMain.l.ListItems.Count
 
   .AddNew
  
  If frmMain.l.ListItems(i).ListSubItems(2).Text <> "" Then .fields("filename") = frmMain.l.ListItems(i).ListSubItems(2).Text
  If frmMain.l.ListItems(i).ListSubItems(3).Text <> "" Then .fields("Path") = frmMain.l.ListItems(i).ListSubItems(3).Text
  If frmMain.l.ListItems(i).ListSubItems(4).Text <> "" Then .fields("artist") = frmMain.l.ListItems(i).ListSubItems(4).Text
  If frmMain.l.ListItems(i).ListSubItems(5).Text <> "" Then .fields("title") = frmMain.l.ListItems(i).ListSubItems(5).Text
  If frmMain.l.ListItems(i).ListSubItems(6).Text <> "" Then .fields("artist") = frmMain.l.ListItems(i).ListSubItems(6).Text
  If frmMain.l.ListItems(i).ListSubItems(7).Text <> "" Then .fields("album") = frmMain.l.ListItems(i).ListSubItems(7).Text
  If frmMain.l.ListItems(i).ListSubItems(8).Text <> "" Then .fields("year") = frmMain.l.ListItems(i).ListSubItems(8).Text
  If frmMain.l.ListItems(i).ListSubItems(9).Text <> "" Then .fields("genre") = frmMain.l.ListItems(i).ListSubItems(9).Text
  If frmMain.l.ListItems(i).ListSubItems(10).Text <> "" Then .fields("length") = frmMain.l.ListItems(i).ListSubItems(10).Text
  If frmMain.l.ListItems(i).ListSubItems(11).Text <> "" Then .fields("size") = frmMain.l.ListItems(i).ListSubItems(11).Text
  
  .fields("Id") = .RecordCount
  .Update
  
  frmNewCollection.p.Value = frmNewCollection.p.Value + 1
  DoEvents
 
 Next i
 End With
 
 db.Close
End Sub



' Add data from an existing REX file to the Music Collection
Public Sub ReadREXFormat(sFilename As String)
    mVariables.musicDataBase = sFilename
    On Error Resume Next
    Dim oldStatusBarMessage As String
    Dim i As Long
    Dim db As Database
    Dim dbrecordset As Recordset
    Dim td As TableDef
 
    Set db = DBEngine.Workspaces(0).OpenDatabase(mVariables.musicDataBase)
    Set dbrecordset = db.OpenRecordset("MusicCollection")
 
    If dbrecordset.RecordCount > 0 Then
        dbrecordset.MoveFirst
    End If
 
    With dbrecordset
        Do While .EOF = False
              
            frmMain.l.ListItems().add = " "
            frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , ""
            frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , .fields("filename")
            frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , .fields("Path")
            frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , .fields("artist")
            frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , .fields("title")
            frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , .fields("artist")
            frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , .fields("album")
            frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , .fields("year")
            frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , .fields("genre")
            frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , .fields("length")
            frmMain.l.ListItems(frmMain.l.ListItems.Count).ListSubItems.add , , .fields("size")
    
            .MoveNext
            .Update
        Loop
    End With
    db.Close
End Sub




