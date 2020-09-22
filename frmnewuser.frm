VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNewCollection 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5535
   ClientLeft      =   30
   ClientTop       =   -15
   ClientWidth     =   7455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmnewuser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar p 
      Height          =   255
      Left            =   360
      TabIndex        =   46
      Top             =   5160
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Max             =   13
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   45
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton prev 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   372
      Left            =   6000
      TabIndex        =   22
      Top             =   4440
      Width           =   612
   End
   Begin VB.PictureBox picLogo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   5292
      Left            =   0
      ScaleHeight     =   5295
      ScaleWidth      =   255
      TabIndex        =   43
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton n 
      Caption         =   ">"
      Height          =   372
      Left            =   6720
      TabIndex        =   21
      Top             =   4440
      Width           =   612
   End
   Begin VB.Frame f3 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3732
      Left            =   3240
      TabIndex        =   30
      Top             =   600
      Visible         =   0   'False
      Width           =   4092
      Begin VB.TextBox faxnumber 
         Height          =   288
         Left            =   1320
         TabIndex        =   14
         Top             =   2640
         Width           =   2772
      End
      Begin VB.TextBox mobilephone 
         Height          =   288
         Left            =   1320
         TabIndex        =   13
         Top             =   2160
         Width           =   2772
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Create DataBase"
         Default         =   -1  'True
         Height          =   372
         Left            =   1320
         TabIndex        =   34
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox workextension 
         Height          =   288
         Left            =   1320
         TabIndex        =   12
         Top             =   1440
         Width           =   2772
      End
      Begin VB.TextBox workphone 
         Height          =   288
         Left            =   1320
         TabIndex        =   11
         Top             =   960
         Width           =   2772
      End
      Begin VB.TextBox title 
         Height          =   288
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   3132
      End
      Begin VB.Label Label14 
         Caption         =   "Fax Number"
         Height          =   252
         Left            =   0
         TabIndex        =   36
         Top             =   2640
         Width           =   1332
      End
      Begin VB.Label m 
         Caption         =   "Mobile Phone"
         Height          =   252
         Left            =   0
         TabIndex        =   35
         Top             =   2160
         Width           =   1212
      End
      Begin VB.Label Label13 
         Caption         =   "Work Extension"
         Height          =   252
         Left            =   0
         TabIndex        =   33
         Top             =   1440
         Width           =   1212
      End
      Begin VB.Label Label12 
         Caption         =   "Workphone"
         Height          =   252
         Left            =   0
         TabIndex        =   32
         Top             =   960
         Width           =   972
      End
      Begin VB.Label Label11 
         Caption         =   "Title"
         Height          =   252
         Left            =   0
         TabIndex        =   31
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.Frame f2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3852
      Left            =   3240
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   4092
      Begin VB.TextBox country 
         Height          =   288
         Left            =   960
         TabIndex        =   9
         Top             =   3120
         Width           =   3132
      End
      Begin VB.TextBox state 
         Height          =   288
         Left            =   960
         TabIndex        =   8
         Top             =   2640
         Width           =   1812
      End
      Begin VB.TextBox city 
         Height          =   288
         Left            =   960
         TabIndex        =   7
         Top             =   2160
         Width           =   3132
      End
      Begin VB.TextBox postalcode 
         Height          =   288
         Left            =   960
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1680
         Width           =   1812
      End
      Begin VB.TextBox company 
         Height          =   288
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   3132
      End
      Begin RichTextLib.RichTextBox address 
         Height          =   732
         Left            =   960
         TabIndex        =   5
         Top             =   720
         Width           =   3132
         _ExtentX        =   5530
         _ExtentY        =   1296
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmnewuser.frx":0442
      End
      Begin VB.Label Label10 
         Caption         =   "Country"
         Height          =   252
         Left            =   0
         TabIndex        =   29
         Top             =   3120
         Width           =   972
      End
      Begin VB.Label Label9 
         Caption         =   "State "
         Height          =   252
         Left            =   0
         TabIndex        =   28
         Top             =   2640
         Width           =   972
      End
      Begin VB.Label Label8 
         Caption         =   "City"
         Height          =   252
         Left            =   0
         TabIndex        =   27
         Top             =   2160
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "Postal code"
         Height          =   252
         Left            =   0
         TabIndex        =   26
         Top             =   1680
         Width           =   1452
      End
      Begin VB.Label Label3 
         Caption         =   "Company"
         Height          =   372
         Left            =   0
         TabIndex        =   25
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label7 
         Caption         =   "Address"
         Height          =   252
         Left            =   0
         TabIndex        =   24
         Top             =   720
         Width           =   852
      End
   End
   Begin VB.Frame f1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3732
      Left            =   3240
      TabIndex        =   15
      Top             =   600
      Width           =   4212
      Begin VB.ComboBox dear 
         Height          =   288
         ItemData        =   "frmnewuser.frx":04C4
         Left            =   3120
         List            =   "frmnewuser.frx":04CE
         TabIndex        =   3
         Top             =   1680
         Width           =   972
      End
      Begin VB.TextBox lastname 
         Height          =   288
         Left            =   960
         TabIndex        =   1
         Top             =   720
         Width           =   3132
      End
      Begin VB.TextBox firstname 
         Height          =   288
         Left            =   960
         MaxLength       =   75
         TabIndex        =   0
         Top             =   240
         Width           =   3132
      End
      Begin VB.TextBox socialnumber 
         Height          =   288
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1200
         Width           =   972
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Dear"
         Height          =   252
         Left            =   1560
         TabIndex        =   20
         Top             =   1680
         Width           =   1452
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Social Security Number :"
         Height          =   252
         Left            =   1320
         TabIndex        =   19
         Top             =   1200
         Width           =   1692
      End
      Begin VB.Label Label1 
         Caption         =   "First Name"
         Height          =   252
         Left            =   0
         TabIndex        =   17
         Top             =   240
         Width           =   1332
      End
      Begin VB.Label Label2 
         Caption         =   "Last Name"
         Height          =   252
         Left            =   0
         TabIndex        =   16
         Top             =   720
         Width           =   1332
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3240
      Top             =   3480
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.PictureBox picNav 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   216
      Index           =   1
      Left            =   240
      ScaleHeight     =   210
      ScaleWidth      =   7095
      TabIndex        =   44
      Top             =   4920
      Width           =   7095
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   108
         X2              =   8760
         Y1              =   24
         Y2              =   24
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   7
         X1              =   108
         X2              =   8760
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame fX 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3732
      Left            =   3240
      TabIndex        =   37
      Top             =   600
      Visible         =   0   'False
      Width           =   4092
      Begin VB.CommandButton Command2 
         Caption         =   "&GO"
         Height          =   372
         Left            =   1320
         TabIndex        =   42
         Top             =   2760
         Width           =   1452
      End
      Begin VB.OptionButton Option3 
         Caption         =   "&Edit your profile"
         Height          =   252
         Left            =   600
         TabIndex        =   41
         Top             =   2040
         Width           =   3132
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Create another user profile"
         Height          =   252
         Left            =   600
         TabIndex        =   40
         Top             =   1560
         Width           =   3252
      End
      Begin VB.OptionButton Option1 
         Caption         =   "&Start working with your new user profile"
         Height          =   252
         Left            =   600
         TabIndex        =   39
         Top             =   1080
         Value           =   -1  'True
         Width           =   3252
      End
      Begin VB.Label Label15 
         Caption         =   "Your database was created succesfully. Please select an option below and then click the 'GO' button."
         Height          =   612
         Left            =   0
         TabIndex        =   38
         Top             =   240
         Width           =   4092
      End
   End
   Begin VB.Label cap 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000C&
      Caption         =   "Personal information "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   120
      Width           =   4095
   End
   Begin VB.Image picture1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4695
      Left            =   360
      Picture         =   "frmnewuser.frx":04DD
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2730
   End
End
Attribute VB_Name = "frmNewCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cl As New cLogo
Private Sub drawlogo()
    cl.DrawingObject = picLogo
    cl.Caption = "   New Music Collection"
End Sub


Private Sub ReDrawLogo()
    On Error Resume Next
    picLogo.Height = Me.ScaleHeight
    On Error GoTo 0
    cl.Draw
End Sub




Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

 Call checkdata
 
End Sub



Private Sub NewDatabase_Click()

 
       
End Sub

Sub CreateDataBase()
    n.Visible = False
    prev.Visible = False


  'Búa til skráarnafn fyrir Notandann.
  Dim sNewName As String
  cd.DialogTitle = "Save User Information" '"Save New User"
  cd.FilterIndex = 1
  cd.Filter = "Rex music files (*.rex)|*.rex|All Files (*.*)|*.*"
  cd.InitDir = App.Path & "\Data"
  cd.Filename = vbNullString
  cd.flags = FileOpenConstants.cdlOFNOverwritePrompt + FileOpenConstants.cdlOFNHideReadOnly
  cd.ShowSave
  DoEvents
  
  If Len(cd.Filename) > 0 Then
    sNewName = cd.Filename
    If InStr(sNewName, ".") = 0 Then
      'add an extension if the user didn't supply one
      sNewName = sNewName & ".rex"
    End If
    If Dir(sNewName) <> vbNullString Then
      Kill sNewName
    End If
  Else
    Command1.Enabled = True
    Exit Sub
  End If
  If Len(sNewName) = 0 Then
    Command1.Enabled = True
    Exit Sub
  End If


    ' Breytur skilgreindar.
    Dim db As Database
    Dim td As TableDef, TempTd As TableDef
    Dim fields(100) As field, indexfield As field
    Dim dbindex As Index
    Dim dbrecordset As Recordset
    Dim t As Integer ' Hverrar tegunar er svæðið.
    
    ' Innfærðar upplýsingar fyrir hendi, má stofna gagnagrunn.
    Command1.Enabled = False
    p.Visible = True

    If frmMain.l.ListItems.Count > 0 Then p.Max = p.Max + frmMain.l.ListItems.Count
    p.Value = 0

    
    ' Það er í lagi með gefið skráarnafn.
    Set db = DBEngine.Workspaces(0).CreateDataBase(sNewName, dbLangGeneral, dbVersion30)
    p.Value = 1


    ' User Information taflan.
    Set td = db.CreateTableDef("User Information")
     Dim i As Long
     Dim field As String
    For i = 0 To 36
     If i = 0 Then field = "FirstName"
     If i = 1 Then field = "LastName"
     If i = 2 Then field = "ID"
     If i = 3 Then field = "Address"
     If i = 4 Then field = "Telephone"
     If i = 5 Then field = "Email"
     If i = 6 Then field = "Website"
     If i = 7 Then field = "Phone"
     If i = 8 Then field = "Fax"
     If i = 9 Then field = "PostBox"
     If i = 10 Then field = "Mobile"
     If i = 11 Then field = "City"
     If i = 12 Then field = "PostCode"
     If i = 13 Then field = "Dear"
     If i = 14 Then field = "Company"
     If i = 15 Then field = "State"
     If i = 16 Then field = "country"
     If i = 17 Then field = "Title"
     If i = 18 Then field = "WorkPhone"
     If i = 19 Then field = "WorkExtension"
     If i = 20 Then field = "Mobilephone"
     If i = 21 Then field = "Rank"
     If i = 22 Then field = "NetWorkID"
     If i = 23 Then field = "CreditCardNumber"
     If i = 24 Then field = "Expires"
     If i = 25 Then field = "CardHolder"
     If i = 26 Then field = "UsePassword"
     If i = 27 Then field = "Organization" ' Boolean. Used with e-mails.
     If i = 28 Then field = "UseSignature"
     If i = 29 Then field = "Signature" ' MEMO
     If i = 30 Then field = "WebStartupPage"
     If i = 31 Then field = "CompanyAddress" ' MEMO
     If i = 32 Then field = "CompanyPhone"
     If i = 33 Then field = "CompanyFax"
     If i = 34 Then field = "CompanyEmails" ' MEMO
     If i = 35 Then field = "CompanyWebsite"
     If i = 36 Then field = "NMCompany" ' Name of the Network Marketing company."
     
  
     If i = 26 Or i = 28 Then
      Set fields(i) = td.CreateField(field, dbBoolean)
      p.Value = 2
     ElseIf i = 29 Or i = 31 Or i = 34 Then
      Set fields(i) = td.CreateField(field, dbMemo)
     ElseIf i <> 3 Then
      Set fields(i) = td.CreateField(field, dbText)
     Else
      Set fields(i) = td.CreateField(field, dbMemo)
     End If
     
     td.fields.Append fields(i)
    Next i

    Set dbindex = td.CreateIndex("ID" & "index")
    Set indexfield = dbindex.CreateField("ID")
    dbindex.fields.Append indexfield
    td.Indexes.Append dbindex
    db.TableDefs.Append td
    p.Value = 3 ' Statusbar
    
    ' NetWork taflan
    Dim tnetwork As TableDef
    Set tnetwork = db.CreateTableDef("Preferences")
      
 
    For i = 0 To 14
     If i = 0 Then
      field = "IDNumber"
      t = 2
     ElseIf i = 1 Then
      field = "Name"
      t = 1
     ElseIf i = 2 Then
      field = "Rank"
      t = 2
     ElseIf i = 3 Then
      field = "Phone"
      t = 1
     ElseIf i = 4 Then
      field = "Email"
      t = 1
     ElseIf i = 5 Then
      field = "Status"
      t = 2
     ElseIf i = 6 Then
      field = "Info"
      t = 3
     ElseIf i = 7 Then
      field = "TeamMember"
      t = 4
     ElseIf i = 8 Then
      field = "TeamName"
      t = 1
     ElseIf i = 9 Then
      field = "Address"
      t = 1
     ElseIf i = 10 Then
      field = "Dear"
      t = 1
     ElseIf i = 11 Then
      field = "Mobile"
      t = 1
     ElseIf i = 12 Then
      field = "Fax"
      t = 1
     ElseIf i = 13 Then
      field = "WebSite"
      t = 1
     ElseIf i = 14 Then
      field = "Newsletter"
      t = 1
     ElseIf i = 15 Then
      field = "Picture"
      t = 5
     End If
     
  
    If t = 1 Then
     Set fields(i) = tnetwork.CreateField(field, dbText)
    ElseIf t = 2 Then
     Set fields(i) = tnetwork.CreateField(field, dbLong)
    ElseIf t = 3 Then
     Set fields(i) = tnetwork.CreateField(field, dbMemo)
    ElseIf t = 4 Then
     Set fields(i) = tnetwork.CreateField(field, dbBoolean)
    ElseIf t = 5 Then
     Set fields(i) = tnetwork.CreateField(field, dbLongBinary)
    End If
    
     tnetwork.fields.Append fields(i)
    Next i

    Set dbindex = tnetwork.CreateIndex("IDNumber" & "index")
    Set indexfield = dbindex.CreateField("IDNumber")
    dbindex.fields.Append indexfield
    tnetwork.Indexes.Append dbindex
    db.TableDefs.Append tnetwork
    p.Value = 4 ' Statusbar

    ' Tekjutaflan
    Dim tcommision As TableDef
    Set tcommision = db.CreateTableDef("Commision")
      
 
    For i = 0 To 6
     If i = 0 Then field = "Year"
     If i = 1 Then field = "Week"
     If i = 2 Then field = "Commision"
     If i = 3 Then field = "LPts"
     If i = 4 Then field = "RPts"
     If i = 5 Then field = "MAX"
     If i = 6 Then field = "Info"
  
     Set fields(i) = tcommision.CreateField(field, dbText)
     tcommision.fields.Append fields(i)
    Next i

    Set dbindex = tcommision.CreateIndex("Year" & "index")
    Set indexfield = dbindex.CreateField("Year")
    dbindex.fields.Append indexfield
    tcommision.Indexes.Append dbindex
    db.TableDefs.Append tcommision
    p.Value = 5 ' Statusbar


    ' Distributors
    Dim tcontacts As TableDef
    Set tcontacts = db.CreateTableDef("MusicCollection")
      
 
    For i = 0 To 37
     If i = 0 Then field = "ID"
     If i = 1 Then field = "P"
     If i = 2 Then field = "R"
     If i = 3 Then field = "Filename"
     If i = 4 Then field = "Path"
     If i = 5 Then field = "Artist"
     If i = 6 Then field = "Title"
     If i = 7 Then field = "Album"
     If i = 8 Then field = "Year"
     If i = 9 Then field = "Genre"
     If i = 10 Then field = "Length"
     If i = 11 Then field = "Size"
     If i = 12 Then field = "Label"
     If i = 13 Then field = "Comments"
     If i = 14 Then field = "Bitrate"
     If i = 15 Then field = "Frequency"
     If i = 16 Then field = "Duration"
     If i = 17 Then field = "Version"
     If i = 18 Then field = "Layer"
     If i = 19 Then field = "Original"
     If i = 20 Then field = "Empasis"
     If i = 21 Then field = "Copyright"
     If i = 22 Then field = "Mode"
     If i = 23 Then field = "Private"
     If i = 24 Then field = "Padding"
     If i = 25 Then field = "CRC"
     If i = 26 Then field = "Payment"
     If i = 27 Then field = "CreditCard"
     If i = 28 Then field = "Valid"
     If i = 29 Then field = "OtherPayment"
     If i = 30 Then field = "dLAccessed"
     If i = 31 Then field = "dModified"
     If i = 32 Then field = "dCreated"
     If i = 33 Then field = "c3"
     If i = 34 Then field = "c4"
     If i = 35 Then field = "v1"
     If i = 36 Then field = "v2"
     If i = 37 Then field = "PaymentNotes"
     If i = 38 Then field = "Spouse"
     If i = 39 Then field = "Children"
     If i = 40 Then field = "Cover"
     If i = 41 Then field = "Preview"
     
     
     
    If i = 39 Then
     Set fields(i) = tcontacts.CreateField(field, dbMemo)
    ElseIf i <> 30 Then
     Set fields(i) = tcontacts.CreateField(field, dbText)
    ElseIf i = 30 Then
     Set fields(i) = tcontacts.CreateField(field, dbLong)
    End If
     tcontacts.fields.Append fields(i)
    Next i

    Set dbindex = tcontacts.CreateIndex("ID" & "index")
    Set indexfield = dbindex.CreateField("ID")
    dbindex.fields.Append indexfield
    tcontacts.Indexes.Append dbindex
    db.TableDefs.Append tcontacts
    p.Value = 6 ' Statusbar
 
 
 
 
    ' Póstkassinn
    Dim tmail As TableDef
    Set tmail = db.CreateTableDef("Mail")
      
 
    For i = 0 To 3
     If i = 0 Then field = "IDNumber"
     If i = 1 Then field = "Inbox"
     If i = 2 Then field = "OutBox"
     If i = 3 Then field = "Drafts"
     If i = 4 Then field = "Deleted"
 
     Set fields(i) = tmail.CreateField(field, dbText)
     tmail.fields.Append fields(i)
    Next i

    Set dbindex = tmail.CreateIndex("IDNumber" & "index")
    Set indexfield = dbindex.CreateField("IDNumber")
    dbindex.fields.Append indexfield
    tmail.Indexes.Append dbindex
    db.TableDefs.Append tmail
    p.Value = 7 ' Statusbar
 
 
 
 
 
 
 
    ' Vörur fyrirtækisins.
    Dim tproducts As TableDef
    Set tproducts = db.CreateTableDef("PlayList")
      
 
    For i = 0 To 4
     If i = 0 Then
      field = "pNumber"
      t = 1
     ElseIf i = 1 Then
      field = "pIcon"
      t = 2
     ElseIf i = 2 Then
      field = "pPts"
      t = 1
     ElseIf i = 3 Then
      field = "pDescribtion"
      t = 3
     ElseIf i = 4 Then
      field = "pImage"
      t = 2
     End If
 
     If t = 1 Then
      Set fields(i) = tproducts.CreateField(field, dbText)
     ElseIf t = 2 Then
      Set fields(i) = tproducts.CreateField(field, dbLongBinary)
     ElseIf t = 3 Then
      Set fields(i) = tproducts.CreateField(field, dbMemo)
     End If
     
     tproducts.fields.Append fields(i)
    Next i

    Set dbindex = tproducts.CreateIndex("pNumber" & "index")
    Set indexfield = dbindex.CreateField("pNumber")
    dbindex.fields.Append indexfield
    tproducts.Indexes.Append dbindex
    db.TableDefs.Append tproducts
    p.Value = 8 ' Statusbar
 
 
 
    ' Taflan sem inniheldur upplýsingar um félagastarfsemi, Team.
    Dim tTeam As TableDef
    Set tTeam = db.CreateTableDef("Teams")
      
 
    For i = 0 To 11
     If i = 0 Then
      field = "Name"
      t = 1
     ElseIf i = 1 Then
      field = "Address"
      t = 3
     ElseIf i = 2 Then
      field = "SocialID"
      t = 1
     ElseIf i = 3 Then
      field = "Founded"
      t = 1
     ElseIf i = 4 Then
      field = "Phone"
      t = 1
     ElseIf i = 5 Then
      field = "email"
      t = 1
     ElseIf i = 6 Then
      field = "Fax"
      t = 1
     ElseIf i = 7 Then
      field = "Website"
      t = 1
     ElseIf i = 8 Then
      field = "TeamImage"
      t = 2
     ElseIf i = 9 Then
      field = "President"
      t = 1
     ElseIf i = 10 Then
      field = "IDNumber"
      t = 1
     ElseIf i = 11 Then
      field = "TeamKey"
      t = 4
     End If
 
     If t = 1 Then
      Set fields(i) = tTeam.CreateField(field, dbText)
     ElseIf t = 2 Then
      Set fields(i) = tTeam.CreateField(field, dbLongBinary)
     ElseIf t = 3 Then
      Set fields(i) = tTeam.CreateField(field, dbMemo)
     ElseIf t = 4 Then
      Set fields(i) = tTeam.CreateField(field, dbLong)
     End If
     
     tTeam.fields.Append fields(i)
    Next i

    Set dbindex = tTeam.CreateIndex("TeamKey" & "index")
    Set indexfield = dbindex.CreateField("TeamKey")
    dbindex.fields.Append indexfield
    tTeam.Indexes.Append dbindex
    db.TableDefs.Append tTeam
    p.Value = 9 ' Statusbar
 
 
 
 
  ' Taflan sem inniheldur upplýsingar um sjálfan gagnagrunninn.
    Dim tDBase As TableDef
    Set tDBase = db.CreateTableDef("Find Collection")
      
 
    For i = 0 To 13
     If i = 0 Then
      field = "CrDate"
      t = 1
     ElseIf i = 1 Then
      field = "CrTime"
      t = 1
     ElseIf i = 2 Then
      field = "Original location"
      t = 1
     ElseIf i = 3 Then
      field = "lModified"
      t = 1
     ElseIf i = 4 Then
      field = "lAccessed"
      t = 1
     ElseIf i = 5 Then
      field = "MS-DOS name"
      t = 1
     ElseIf i = 6 Then
      field = "BackUp"
      t = 1
     ElseIf i = 7 Then
      field = "SecurityLevel"
      t = 1
     ElseIf i = 8 Then
      field = "Addition information"
      t = 3
     ElseIf i = 9 Then
      field = "History"
      t = 3
     ElseIf i = 10 Then
      field = "Errors"
      t = 3
     ElseIf i = 11 Then
      field = "Username"
      t = 1
     ElseIf i = 12 Then
      field = "Password"
      t = 1
     ElseIf i = 13 Then
      field = "DataBaseKey"
     End If
 
     If t = 1 Then
      Set fields(i) = tDBase.CreateField(field, dbText)
     ElseIf t = 2 Then
      Set fields(i) = tDBase.CreateField(field, dbLongBinary)
     ElseIf t = 3 Then
      Set fields(i) = tDBase.CreateField(field, dbMemo)
     ElseIf t = 4 Then
      Set fields(i) = tDBase.CreateField(field, dbLong)
     End If
     
     tDBase.fields.Append fields(i)
    Next i

    Set dbindex = tDBase.CreateIndex("DatabaseKey" & "index")
    Set indexfield = dbindex.CreateField("DataBaseKey")
    dbindex.fields.Append indexfield
    tDBase.Indexes.Append dbindex
    db.TableDefs.Append tDBase
    p.Value = 10 ' Statusbar
 
 
 
 
 
    ' Upplýsingar sem notandinn hefur fært inn, skráðar.
   ' Hvaða töflu á að skrá gögn í.
   Set dbrecordset = db.OpenRecordset("User Information", dbOpenTable)
   p.Value = 11 ' Statusbar
 
   ' Ný færsla
   dbrecordset.AddNew
   
   ' Gögnin skráð.
   If firstname.Text <> "" Then dbrecordset.fields("FirstName") = firstname.Text
   If lastname.Text <> "" Then dbrecordset.fields("LastName") = lastname.Text
   If socialnumber.Text <> "" Then dbrecordset.fields("ID") = socialnumber.Text
   If address.Text <> "" Then dbrecordset.fields("Address") = address.Text
   If company.Text <> "" Then dbrecordset.fields("Company") = company.Text
   If dear.Text <> "" Then dbrecordset.fields("Dear") = dear.Text
   If address.Text <> "" Then dbrecordset.fields("Address") = address.Text
   If postalcode.Text <> "" Then dbrecordset.fields("PostCode") = postalcode.Text
   If state.Text <> "" Then dbrecordset.fields("State") = state.Text
   If country.Text <> "" Then dbrecordset.fields("Country") = country.Text
   If title.Text <> "" Then dbrecordset.fields("Title") = title.Text
   If workphone.Text <> "" Then dbrecordset.fields("Workphone") = workphone.Text
   If workextension.Text <> "" Then dbrecordset.fields("Workextension") = workextension.Text
   If mobilephone.Text <> "" Then dbrecordset.fields("mobilephone") = mobilephone.Text
   If faxnumber.Text <> "" Then dbrecordset.fields("fax") = faxnumber.Text
   If city.Text <> "" Then dbrecordset.fields("city") = city.Text
   dbrecordset.fields("UsePassword") = False



   
   
   
   dbrecordset.Update
   
   p.Value = 12 ' Statusbar
   
   
   
  ' Taflan sem inniheldur upplýsingar um Tekjustöðvarnar.
    Dim tIncomeC As TableDef
    Set tIncomeC = db.CreateTableDef("IncomeCenters")
      
 
    For i = 0 To 12
     If i = 0 Then
      field = "IncomeCenter"
      t = 1
     ElseIf i = 1 Then
      field = "Extension"
      t = 1
     ElseIf i = 2 Then
      field = "CrDate"
      t = 1
     ElseIf i = 3 Then
      field = "CrTime"
      t = 1
     ElseIf i = 4 Then
      field = "LastPaydTo"
      t = 1
     ElseIf i = 5 Then
      field = "LastModified"
      t = 1
     ElseIf i = 6 Then
      field = "Owner"
      t = 1
     ElseIf i = 7 Then
      field = "SecurityLevel"
      t = 1
     ElseIf i = 8 Then
      field = "Addition information"
      t = 3
     ElseIf i = 9 Then
      field = "Errors"
      t = 3
     ElseIf i = 10 Then
      field = "Username"
      t = 1
     ElseIf i = 11 Then
      field = "Password"
      t = 1
     ElseIf i = 12 Then
      field = "DataBaseKey"
      t = 1
     End If
 
     If t = 1 Then
      Set fields(i) = tIncomeC.CreateField(field, dbText)
     ElseIf t = 2 Then
      Set fields(i) = tIncomeC.CreateField(field, dbLongBinary)
     ElseIf t = 3 Then
      Set fields(i) = tIncomeC.CreateField(field, dbMemo)
     ElseIf t = 4 Then
      Set fields(i) = tIncomeC.CreateField(field, dbLong)
     End If
     
     tIncomeC.fields.Append fields(i)
    Next i

    Set dbindex = tIncomeC.CreateIndex("DatabaseKey" & "index")
    Set indexfield = dbindex.CreateField("DataBaseKey")
    dbindex.fields.Append indexfield
    tIncomeC.Indexes.Append dbindex
    db.TableDefs.Append tIncomeC
    p.Value = 13 ' Statusbar
 
   
   
   db.Close
   
   
  ' mVariables.MusicCollectionDatabase("Svenni") = cd.FileName  ' Vistaðar upplýsingar um gagnagrunn notanda.
   'frmMain.Caption = mVariables.MusicCollectionDatabase.FileName
   'frmNetWorker.StatusBar1.Panels(1).Text
   mVariables.musicDataBase = cd.Filename
   If frmMain.l.ListItems.Count > 0 Then
        Call mDataBase.AddCollection
    End If
   
   Command1.Enabled = True
   p.Visible = False
   
   ' Lokaákvörðun notanda.
   fX.Visible = True
   prev.Enabled = False
   n.Enabled = False
   n.Visible = True
   prev.Visible = True

  
   
End Sub



Sub checkdata()
 Dim msg, style, title, help, ctxt, response, MyString
 title = ""
 If socialnumber.Text = "" Then msg = "Social Security number (ID), missing."
 If firstname.Text = "" Then msg = "User Name missing"

 
 
 
 
 If msg <> "" Then
  title = "Error"   ' Define message.
  style = vbOKOnly + vbCritical + vbDefaultButton1   ' Define buttons.
  ctxt = 1000   ' Define topic
  response = MsgBox(msg, style, title, help, ctxt)
  Command1.Enabled = True
  Exit Sub ' Hætt við að vista.
 Else
  Call CreateDataBase ' Má vista.
 End If
End Sub




Private Sub Command2_Click()

' If Option1.Value = True Then
'  Call proc.en("true")
'  Unload Me
'  frmTaskWizard.Show vbModal
' ElseIf Option2.Value = True Then
'  Call add
' ElseIf Option3.Value = True Then
'  Call edit
' End If


End Sub

Private Sub Form_Load()
  Call drawlogo
 On Error GoTo e:
  picture1.Picture = LoadPicture(App.Path & "\Images\AddUser.jpg")
e:
End Sub

Private Sub Form_Resize()
 Call ReDrawLogo
End Sub

Private Sub n_Click()
 If f1.Visible = True Then
  f1.Visible = False
  f2.Visible = True
  cap.Caption = "Address information "
  prev.Enabled = True
 ElseIf f2.Visible = True Then
  f2.Visible = False
  f3.Visible = True
  cap.Caption = "Contact information "
  n.Enabled = False
 End If
 
End Sub



Sub add()
 firstname.Text = ""
 lastname.Text = ""
 socialnumber.Text = ""
 dear.Text = ""
 company.Text = ""
 title.Text = ""
 address.Text = ""
 postalcode.Text = ""
 mobilephone.Text = ""
 workextension.Text = ""
 workphone.Text = ""
 state.Text = ""
 country.Text = ""
 faxnumber.Text = ""
 f1.Visible = True
 n.Enabled = True
 cap.Caption = "Personal information "
 fX.Visible = False
End Sub

Sub edit()
 fX.Visible = False
 n.Enabled = True
 prev.Enabled = False
 f1.Visible = True
 cap.Caption = "Personal information "
 f3.Visible = False
End Sub

Private Sub prev_Click()
 If f2.Visible = True Then
  f2.Visible = False
  f1.Visible = True
  cap.Caption = "Personal information "
  prev.Enabled = False
 ElseIf f3.Visible = True Then
  f3.Visible = False
  f2.Visible = True
  cap.Caption = "Address information "
  n.Enabled = True
 End If
 
 
End Sub
