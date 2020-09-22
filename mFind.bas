Attribute VB_Name = "mFind"
' Routines for findind data in DataBase


Option Explicit



' Find Files in Collection
Public Sub Find()
    Dim i As Long
    With frmMain
        .lf.ListItems.Clear
        .l.Visible = False
        .lf.Visible = True
        For i = 1 To .l.ListItems.Count
            If .tTitle.Text > "" And .tPath.Text > "" Then
                If InStr(UCase(.l.ListItems(i).ListSubItems(2).Text), UCase(.tTitle.Text)) > 0 And InStr(UCase(.l.ListItems(i).ListSubItems(3).Text), UCase(.tPath.Text)) > 0 Then
                    '.lf.AddItem (.l.ListItems(i).ListSubItems(3).Text & "\" & .l.ListItems(i).ListSubItems(2).Text)
                    Call AddToFindGrid(i)
                End If
            ElseIf .tPath.Text = "" And .tTitle.Text > "" Then
                If InStr(UCase(.l.ListItems(i).ListSubItems(2).Text), UCase(.tTitle.Text)) > 0 Then
                    'lSongs.AddItem (.l.ListItems(i).ListSubItems(3).Text & "\" & .l.ListItems(i).ListSubItems(2).Text)
                    Call AddToFindGrid(i)
               End If
           ElseIf .tPath.Text > "" And .tTitle.Text = "" Then
                If InStr(UCase(.l.ListItems(i).ListSubItems(3).Text), UCase(.tPath.Text)) > 0 Then
                    'lSongs.AddItem (.l.ListItems(i).ListSubItems(3).Text & "\" & .l.ListItems(i).ListSubItems(2).Text)
                    Call AddToFindGrid(i)
                End If
           End If
        Next i
    End With
End Sub




' Add a Result to the Find Grid
Private Sub AddToFindGrid(i As Long)
        frmMain.lf.ListItems().Add = " "
        frmMain.lf.ListItems(frmMain.lf.ListItems.Count).ListSubItems.Add , , ""
        frmMain.lf.ListItems(frmMain.lf.ListItems.Count).ListSubItems.Add , , frmMain.l.ListItems(i).ListSubItems(2).Text
        frmMain.lf.ListItems(frmMain.lf.ListItems.Count).ListSubItems.Add , , frmMain.l.ListItems(i).ListSubItems(3).Text
        frmMain.lf.ListItems(frmMain.lf.ListItems.Count).ListSubItems.Add , , "Sveinn R. Sigurðsson"
        frmMain.lf.ListItems(frmMain.lf.ListItems.Count).ListSubItems.Add , , "Perfect Play of Life"
        frmMain.lf.ListItems(frmMain.lf.ListItems.Count).ListSubItems.Add , , "Upgrade"
        frmMain.lf.ListItems(frmMain.lf.ListItems.Count).ListSubItems.Add , , "2000"
        frmMain.lf.ListItems(frmMain.lf.ListItems.Count).ListSubItems.Add , , "4:10"
        frmMain.lf.ListItems(frmMain.lf.ListItems.Count).ListSubItems.Add , , "D:\Best of pop\DUO\ppl.mp3"
        frmMain.lf.ListItems(frmMain.lf.ListItems.Count).ListSubItems.Add , , "4.103 kb"
End Sub



' Find recursive files on hard disk
Public Sub FindFiles(sPath As String)

   Dim FP As FILE_PARAMS
   
   Call DisplayInit
   
   With FP
      .sFileRoot = sPath
      .sFileNameExt = "*.mp3"
      .bRecurse = 1 'TEMPchkRecurse.Value = 1
      .bList = 1 'chkListResults.Value = 0
   End With
   
   Call SearchForFiles(FP)
   Call DisplayResults(FP)
   Call mCollection.AddFiles(frmMain)
   
End Sub

Private Sub Command2_Click()

'   Dim FP As FILE_PARAMS
'
'   Call DisplayInit
'
'   With FP
'      .sFileRoot = frmMain.Text1.Text
'      .sFileNameExt = "*.*"
'      .bRecurse = chkRecurse.Value = 1
'      .bList = chkListResults.Value = 0
'   End With
'
'   Call SearchForFolders(FP)
'   Call DisplayResults(FP)
'
End Sub


Private Sub Command3_Click()

'   Dim FP As FILE_PARAMS
'
'   Call DisplayInit
'
'   With FP
'      .sFileRoot = "c:\"   '"c:\winnt\"
'      .sFileNameExt = "wordpad.exe" ' "notepad.exe"
''   End With
'
'   Call SearchPathForFile(FP)
'   Call DisplayResults(FP)'

End Sub


Private Sub Command4_Click()

'   Dim FP As FILE_PARAMS
'
'   Call DisplayInit
'
'   With FP
'      .sFileRoot = "c:\"
'      .sFileNameExt = "vb6.exe"  '"wordpad.exe"
'   End With
'
'   Call SearchSystemForFile(FP)
'   Call DisplayResults(FP)
   
End Sub


Private Sub DisplayInit()

  'common routine to initialize display
   
   'Text2.Text = "Working ..."
   'Text3.Text = ""
   'Text2.Refresh
   'Text3.Refresh
   frmMain.List1.Clear
   frmMain.List1.Visible = False
   
End Sub


Private Sub DisplayResults(FP As FILE_PARAMS)

  'a common routine to display search results

  'this defaults to show the size and count
  'containing in the FP type members, but if
  'FP.sResult is filled (from the Drive and
  'System search methods), that is shown instead.
  
   'frmMain.Caption = Format$(FP.nFileCount, sFmtResult) & _
                   " (" & FP.sFileNameExt & ")"
                   
   'frmMain.Caption = Format$(FP.nFileSize, sFileSizeBytes)
                                    
   If FP.sResult > "" Then
   
      frmMain.sbrMain.Panels(3).Text = "found:    " & FP.bFound
      frmMain.sbrMain.Panels(1).Text = "location: " & FP.sResult
   
   End If

End Sub


Private Function QualifyPath(sPath As String) As String

  'assures that a passed path ends in a slash
  
   If Right$(sPath, 1) <> "\" Then
         QualifyPath = sPath & "\"
   Else: QualifyPath = sPath
   End If
      
End Function


Function StripItem(startStrg As String) As String

  'Take a string separated by Chr(0)'s,
  'and split off 1 item, and shorten the
  'string so that the next item is ready
  'for removal.
   Dim pos As Integer
   
   pos = InStr(startStrg, Chr$(0))
   
   If pos Then
      StripItem = Mid(startStrg, 1, pos - 1)
      startStrg = Mid(startStrg, pos + 1, Len(startStrg))
   End If
   
End Function


Public Function TrimNull(startstr As String) As String

  'returns the string up to the first
  'null, if present, or the passed string
   Dim pos As Integer
   
   pos = InStr(startstr, Chr$(0))
   
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
  
   TrimNull = startstr
  
End Function


Private Function GetFileInformation(FP As FILE_PARAMS) As Long

  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim nSize As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
      
  'FP.sFileRoot (assigned to sRoot) contains
  'the path to search.
  '
  'FP.sFileNameExt (assigned to sPath) contains
  'the full path and filespec.
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & FP.sFileNameExt
   
  'obtain handle to the first filespec match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then

      Do
      
        'remove trailing nulls
         sTmp = TrimNull(WFD.cFileName)
         
        'Even though this routine uses filespecs,
        '*.* is still valid and will cause the search
        'to return folders as well as files, so a
        'check against folders is still required.
         If Not (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) _
            = FILE_ATTRIBUTE_DIRECTORY Then
      
           'file found, so increase the file count
            FP.nFileCount = FP.nFileCount + 1
            
           'retrieve the size and assign to nSize to
           'be returned at the end of this function call
            nSize = nSize + (WFD.nFileSizeHigh * (MAXDWORD + 1)) + WFD.nFileSizeLow
            
           'add to the list if the flag indicates
            If FP.bList Then frmMain.List1.AddItem sRoot & sTmp
         
         End If
         
      Loop While FindNextFile(hFile, WFD)
      
      
     'close the handle
      hFile = FindClose(hFile)
   
   End If
   
  'return the size of files found
   GetFileInformation = nSize

End Function


Private Function SearchPathForFile(FP As FILE_PARAMS) As Boolean
  
   Dim sResult As String
    
  'pad a return string and search the passed drive
   sResult = Space(MAX_PATH)

  'SearchTreeForFile returns True (1) if found,
  'or False otherwise. If True, sResult holds
  'the full path.
   FP.bFound = SearchTreeForFile(FP.sFileRoot, FP.sFileNameExt, sResult)
       
  'if found, strip the trailing nulls and exit
      If FP.bFound Then
      FP.sResult = LCase$(TrimNull(sResult))
   End If
    
   SearchPathForFile = FP.bFound
    
End Function

Private Function SearchSystemForFile(FP As FILE_PARAMS) As Boolean

   Dim nSize As Long
   Dim sBuffer As String
   Dim currDrive As String
   Dim sResult As String
       
  'retrieve the available drives on the system
   sBuffer = Space$(64)
   nSize = GetLogicalDriveStrings(Len(sBuffer), sBuffer)
   
  'nSize returns the size of the drive string
   If nSize Then
   
     'strip off trailing nulls
      sBuffer = Left$(sBuffer, nSize)
     
     'search each fixed disk drive for the file
      Do Until sBuffer = ""

        'strip off one drive item from sBuffer
         FP.sFileRoot = StripItem(sBuffer)

        'just search the local file system
         If GetDriveType(FP.sFileRoot) = DRIVE_FIXED Then
         
           'this may take a while, so update the
           'display when the search path changes
            frmMain.sbrMain.Panels(1).Text = "Working ... searching drive " & FP.sFileRoot
            
           'pad a return string and search the passed drive
            sResult = Space(MAX_PATH)
      
            FP.bFound = SearchTreeForFile(FP.sFileRoot, FP.sFileNameExt, sResult)
            
           'if found, strip the trailing nulls and exit
            If FP.bFound Then
               FP.sResult = LCase$(TrimNull(sResult))
               Exit Do
            End If
         
         End If
      
      Loop
      
   End If
      
   SearchSystemForFile = FP.bFound

End Function


Private Function SearchForFiles(FP As FILE_PARAMS) As Double

  'local working variables
   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim nSize As Long
   Dim sPath As String
   Dim sRoot As String
   Dim sTmp As String
      
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & "*.*"
   
  'obtain handle to the first match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then
   
     'This is where the method obtains the file
     'list and data for the folder passed.
     '
     'GetFileInformation function returns the size,
     'in bytes, of the files found matching the
     'filespec in the passed folder, so its
     'assigned to nSize. It is not directly assigned
     'to FP.nFileSize because nSize is incremented
     'below if a recursive search was specified.
      nSize = GetFileInformation(FP)
      FP.nFileSize = nSize

      Do
      
        'if the returned item is a folder...
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
            
           '..and the Recurse flag was specified
            If FP.bRecurse Then
            
              'remove trailing nulls
               sTmp = TrimNull(WFD.cFileName)
               
              'and if the folder is not the default
              'self and parent folders...
               If sTmp <> "." And sTmp <> ".." Then
               
                 '..then the item is a real folder, which
                 'may contain other sub folders, so assign
                 'the new folder name to FP.sFileRoot and
                 'recursively call this function again with
                 'the ammended information.
                 '
                 'Since nSize is a local variable, whose value
                 'is both set above as well as returned as the
                 'function call value, the nSize needs to be
                 'added to previous calls in order to maintain accuracy.
                 '
                 'However, because the nFileSize member of
                 'FILE_PARAMS is passed back and forth through
                 'the calls, nSize is simply assigned to it
                 'after the recursive call finishes.
                  FP.sFileRoot = sRoot & sTmp
                  nSize = nSize + SearchForFiles(FP)
                  FP.nFileSize = nSize
                  
               End If
               
            End If
            
         End If
         
     'continue looping until FindNextFile returns
     '0 (no more matches)
      Loop While FindNextFile(hFile, WFD)
      
     'close the find handle
      hFile = FindClose(hFile)
   
   End If
   
  'because this routine is recursive, return
  'the size of matching files
   SearchForFiles = nSize
   
End Function


Private Function SearchForFolders(FP As FILE_PARAMS) As Long

   Dim WFD As WIN32_FIND_DATA
   Dim hFile As Long
   Dim sRoot As String
   Dim sPath As String
   Dim sTmp As String
   Dim nCount As Long
   
   sRoot = QualifyPath(FP.sFileRoot)
   sPath = sRoot & FP.sFileNameExt
   
  'obtain handle to the first match
   hFile = FindFirstFile(sPath, WFD)
   
  'if valid ...
   If hFile <> INVALID_HANDLE_VALUE Then
         
      Do
         
        'We only want folders in this method.
         If (WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
         
           'remove trailing nulls
            sTmp = TrimNull(WFD.cFileName)
         
           'and if not the default system folders
            If sTmp <> "." And sTmp <> ".." Then
            
              'count it and add to the list if the flag indicates
               nCount = nCount + 1
               If FP.bList Then frmMain.List1.AddItem sRoot & sTmp
            
              'if a recursive search was selected, call
              'this method again with a modified root
               If FP.bRecurse Then
               
                  FP.sFileRoot = sRoot & sTmp
                  nCount = nCount + SearchForFolders(FP)
                  
               End If
               
              'this is outside the recurse code in case
              'a single path-search was specified
               FP.nFileCount = nCount
               
            End If
         End If
         
      Loop While FindNextFile(hFile, WFD)
      
     'close the handle
      hFile = FindClose(hFile)
   
   End If

  'since folders are 0-length, return the count instead
   SearchForFolders = nCount
   
End Function

'--end block--'





