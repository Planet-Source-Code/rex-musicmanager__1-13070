Attribute VB_Name = "mPreferences"
Option Explicit



Public Const GWL_STYLE As Long = (-16)
Public Const COLOR_WINDOW As Long = 5
Public Const COLOR_WINDOWTEXT As Long = 8

Public Const TVI_ROOT   As Long = &HFFFF0000
Public Const TVI_FIRST  As Long = &HFFFF0001
Public Const TVI_LAST   As Long = &HFFFF0002
Public Const TVI_SORT   As Long = &HFFFF0003

Public Const TVIF_STATE As Long = &H8

'treeview styles
Public Const TVS_HASLINES As Long = 2
Public Const TVS_FULLROWSELECT As Long = &H1000

'treeview style item states
Public Const TVIS_BOLD  As Long = &H10

Public Const TV_FIRST As Long = &H1100
Public Const TVM_GETNEXTITEM As Long = (TV_FIRST + 10)
Public Const TVM_GETITEM As Long = (TV_FIRST + 12)
Public Const TVM_SETITEM As Long = (TV_FIRST + 13)
Public Const TVM_SETBKCOLOR As Long = (TV_FIRST + 29)
Public Const TVM_SETTEXTCOLOR As Long = (TV_FIRST + 30)
Public Const TVM_GETBKCOLOR As Long = (TV_FIRST + 31)
Public Const TVM_GETTEXTCOLOR As Long = (TV_FIRST + 32)

Public Const TVGN_ROOT                As Long = &H0
Public Const TVGN_NEXT                As Long = &H1
Public Const TVGN_PREVIOUS            As Long = &H2
Public Const TVGN_PARENT              As Long = &H3
Public Const TVGN_CHILD               As Long = &H4
Public Const TVGN_FIRSTVISIBLE        As Long = &H5
Public Const TVGN_NEXTVISIBLE         As Long = &H6
Public Const TVGN_PREVIOUSVISIBLE     As Long = &H7
Public Const TVGN_DROPHILITE          As Long = &H8
Public Const TVGN_CARET               As Long = &H9

Public Type TV_ITEM
   mask As Long
   hItem As Long
   state As Long
   stateMask As Long
   pszText As String
   cchTextMax As Long
   iImage As Long
   iSelectedImage As Long
   cChildren As Long
   lParam As Long
End Type



Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Declare Function GetWindowLong Lib "user32" _
   Alias "GetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" _
   (ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Public Declare Function GetSysColor Lib "user32" _
   (ByVal nIndex As Long) As Long


Public Sub CreatePreferences()
    Dim tIndex As Integer
    Call NewNode("Music Collection", "L")
    Call NewNode("PlayList", "LP")
    Call NewNode("Devices", "D")
    Call NewNode("Find Music", "LF")
    Call NewNode("Visual", "V")
    Call NewNode("Startup", "ST")
    Call NewNode("DataBase", "DB")
    Call CreateChild("Adding", "LPADD", "LP")
    Call CreateChild("Copy selected files to Playlist directory", "LPADDMP3", "LPADD")
    Call CreateChild("Deleting", "LPDEL", "LP")
    Call CreateChild("Delete original files from disk", "LPDELFILE", "LPDEL")
End Sub



Private Sub NewNode(sCaption As String, skey)
On Error Resume Next
    With frmMain
        Dim n As Node
        If .lt.Nodes.Count = 0 Then
            Set n = .lt.Nodes.add()
        Else
            If .lt.SelectedItem.Selected = True Then
                Set n = .lt.Nodes.add(.lt.Nodes(.lt.Nodes.Count), tvwNext, "New")
            Else
                Set n = .lt.Nodes.add()
            End If
        End If
        .lt.Nodes(.lt.Nodes.Count).Text = sCaption
        .lt.Nodes(.lt.Nodes.Count).Key = skey
        .lt.Nodes(.lt.Nodes.Count).BackColor = frmMain.tColor.BackColor
    End With
End Sub



Sub CreateChild(sCaption As String, skey, sBelongsTo As String)
 Dim n As Node
 Dim OldKey As String
 Dim BelongsTo As String

  BelongsTo = sBelongsTo
  Set n = frmMain.lt.Nodes.add(BelongsTo, tvwChild, sCaption)
 
  
  frmMain.lt.Nodes(frmMain.lt.Nodes.Count).Text = sCaption
  frmMain.lt.Nodes(frmMain.lt.Nodes.Count).Key = skey
  frmMain.lt.Nodes(frmMain.lt.Nodes.Count).BackColor = frmMain.tColor.BackColor
End Sub



' Change the TreeViews backgroundcolor, clrref is the color number
Public Sub ChangeBackGroundColor(clrref As Long)
   Dim hwndTV As Long
   Dim style As Long
   
   hwndTV = frmMain.lt.hwnd
   
  'Change the background
   Call SendMessage(hwndTV, TVM_SETBKCOLOR, 0, ByVal clrref)
   
  'reset the treeview style so the
  'tree lines appear properly
   style = GetWindowLong(frmMain.lt.hwnd, GWL_STYLE)
   
  'if the treeview has lines, temporarily
  'remove them so the back repaints to the
  'selected colour, then restore
   If style And TVS_HASLINES Then
      Call SetWindowLong(hwndTV, GWL_STYLE, style Xor TVS_HASLINES)
      Call SetWindowLong(hwndTV, GWL_STYLE, style)
   End If
End Sub





