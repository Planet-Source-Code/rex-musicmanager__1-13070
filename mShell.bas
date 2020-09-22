Attribute VB_Name = "mShell"
Option Explicit

Public Const sFileSizeBytes = "###,###,###,###,###,###,##0 \b\y\t\e\s"
Public Const sFmtResult = "###,###,###,##0 \f\o\u\n\d"
Public Const sFileCount = "###,###,###,##0 f\i\l\e\s\ \f\o\u\n\d"
Public Const sFolderCount = "###,###,###,##0 f\o\l\d\e\r\s \f\o\u\n\d"

Public Const MAXDWORD As Long = &HFFFF
Public Const MAX_PATH As Long = 260
Public Const INVALID_HANDLE_VALUE As Long = -1
Public Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED As Long = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Public Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const FILE_ATTRIBUTE_READONLY As Long = &H1
Public Const FILE_ATTRIBUTE_TEMPORARY As Long = &H100
Public Const FILE_ATTRIBUTE_FLAGS = FILE_ATTRIBUTE_ARCHIVE Or _
                                    FILE_ATTRIBUTE_HIDDEN Or _
                                    FILE_ATTRIBUTE_NORMAL Or _
                                    FILE_ATTRIBUTE_READONLY

Public Const DRIVE_UNKNOWNTYPE As Long = 1
Public Const DRIVE_REMOVABLE As Long = 2
Public Const DRIVE_FIXED As Long = 3
Public Const DRIVE_REMOTE As Long = 4
Public Const DRIVE_CDROM As Long = 5
Public Const DRIVE_RAMDISK As Long = 6

Public Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type

'the custom-made UDT for searching
Public Type FILE_PARAMS
   bRecurse As Boolean     'set True to perform a recursive search
   bList As Boolean        'set True to add results to listbox
   bFound As Boolean       'set only with SearchTreeForFile methods
   sFileRoot As String     'search starting point, ie c:\, c:\winnt\
   sFileNameExt As String  'filenae/filespec to locate, ie *.dll, notepad.exe
   sResult As String       'path to file. Set only with SearchTreeForFile methods
   nFileCount As Long      'total file count matching filespec. Set in FindXXX only
   nFileSize As Double     'total file size matching filespec. Set in FindXXX only
End Type

Public Declare Function FindClose Lib "kernel32" _
  (ByVal hFindFile As Long) As Long
   
Public Declare Function FindFirstFile Lib "kernel32" _
   Alias "FindFirstFileA" _
  (ByVal lpFileName As String, _
   lpFindFileData As WIN32_FIND_DATA) As Long
   
Public Declare Function FindNextFile Lib "kernel32" _
   Alias "FindNextFileA" _
  (ByVal hFindFile As Long, _
   lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function SearchTreeForFile Lib "imagehlp.dll" _
  (ByVal sFileRoot As String, _
   ByVal InputPathName As String, _
   ByVal OutputPathBuffer As String) As Boolean

Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" _
      (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
      
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" _
      (ByVal nDrive As String) As Long
'--end block--'
   
   
   
' Get Volume information from a drive
Declare Function GetVolumeInformation _
    Lib "kernel32" Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, _
    ByVal lpVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, _
    ByVal nFileSystemNameSize As Long) As Long



'Constants for topmost.
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE

' Make forms stay on top
Public Declare Function SetWindowPos _
    Lib "user32" _
   (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, _
    ByVal wFlags As Long) As Long
' End of topMost



' Copy, move, delete files from disk
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
   (lpFileOp As SHFILEOPSTRUCT) As Long


Const FO_MOVE = &H1
Const FO_COPY = &H2
Const FO_DELETE = &H3
Const FO_RENAME = &H4
Const FOF_MULTIDESTFILES = &H1
Const FOF_CONFIRMMOUSE = &H2
Const FOF_SILENT = &H4                      '  don't create progress/report
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_NOCONFIRMATION = &H10             '  Don't prompt the user.
Const FOF_WANTMAPPINGHANDLE = &H20          '  Fill in SHFILEOPSTRUCT.hNameMappings
                                      '  Must be freed using SHFreeNameMappings
Const FOF_ALLOWUNDO = &H40
Const FOF_FILESONLY = &H80                  '  on *.*, do only files - not directories
Const FOF_SIMPLEPROGRESS = &H100            '  means don't show names of files
Const FOF_NOCONFIRMMKDIR = &H200            '  don't confirm making any needed dirs

Const PO_DELETE = &H13           '  printer is being deleted
Const PO_RENAME = &H14           '  printer is being renamed
Const PO_PORTCHANGE = &H20       '  port this printer connected to is being changed
                                '  if this id is set, the strings received by
                                '  the copyhook are a doubly-null terminated
                                '  list of strings.  The first is the printer
                                '  name and the second is the printer port.
Const PO_REN_PORT = &H34         '  PO_RENAME and PO_PORTCHANGE at same time.
 
Public Sub CopyFile(i As Long, dDir As String)
    With frmMain
        If .lt.Nodes.Item("LPADDMP3").Checked = True Then
            Dim lResult As Long, SHF As SHFILEOPSTRUCT
            Dim oldString As String
            SHF.hwnd = frmMain.hwnd
            SHF.wFunc = FO_COPY
            SHF.pFrom = .l.ListItems(i).ListSubItems(3).Text & "\" & .l.ListItems(i).ListSubItems(2).Text
            SHF.pTo = dDir & "\" & .l.ListItems(i).ListSubItems(2).Text
            SHF.fFlags = FOF_FILESONLY + FOF_SILENT
            oldString = .sbrMain.Panels(1).Text
            .sbrMain.Panels(1).Text = "Copying " & .l.ListItems(i).ListSubItems(2).Text & " ..."
            DoEvents
            lResult = SHFileOperation(SHF)
            DoEvents
            .sbrMain.Panels(1).Text = oldString
            If lResult Then
                MsgBox "Could not copy file.", vbCritical + vbOKOnly, "Error"
            End If
        End If
    End With
End Sub





