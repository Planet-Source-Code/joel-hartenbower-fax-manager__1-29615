VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Search"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton SearchBy 
      Caption         =   "Fax Number"
      Height          =   255
      Index           =   3
      Left            =   2453
      TabIndex        =   8
      Top             =   720
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton SearchBy 
      Caption         =   "Last Name"
      Height          =   255
      Index           =   1
      Left            =   2453
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton SearchBy 
      Caption         =   "First Name"
      Height          =   255
      Index           =   0
      Left            =   293
      TabIndex        =   6
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton SearchBy 
      Caption         =   "Company"
      Height          =   255
      Index           =   2
      Left            =   293
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Add New"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      Height          =   375
      Left            =   2640
      MaskColor       =   &H00808080&
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc datFaxDB 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=FaxManagerDB"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "FaxManagerDB"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "datFaxDB"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1560
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
      Caption         =   "Creating Search Routines, Please Wait..."
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Search String:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim strSearch As String
    Dim strCriteria As String
    Dim Response As String
    Dim strField As String
    
    frmSearch.Height = 2250
    On Error Resume Next
    ProgressBar1.Visible = False
    lblPercent.Visible = False
    lblWait.Visible = True
    frmSearch.Refresh
        
    strSearch = Me.txtSearch
    If Me.SearchBy(0).Value = True Then
        strField = "FirstName"
    End If
    If Me.SearchBy(1).Value = True Then
        strField = "LastName"
    End If
    If Me.SearchBy(2).Value = True Then
        strField = "Company"
    End If
    If Me.SearchBy(3).Value = True Then
        strField = "FaxNumber"
    End If
    
    gSQL = "SELECT * FROM [Phone Book Table]"
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh
    
    datFaxDB.Recordset.Filter = strField & " LIKE '" & strSearch & "'"
    If datFaxDB.Recordset.RecordCount > 0 Then
        Unload frmResultsEdit
        frmResultsEdit.Show 0, frmSearch
        Exit Sub
    Else
        strCriteria = strField & " = " & Chr(34) & strSearch & Chr(34)
        rstDelSearch.MoveFirst
        rstDelSearch.FindFirst strCriteria
        Response = MsgBox("That record was already deleted or does not exsist.", vbInformation, "Record Information")
        frmSearch.Height = 1950
        frmSearch.Refresh
        Unload frmResultsEdit
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    gRecordID = 0
    gSearch = 0
    Unload Me
    frmAddEdit.Show 0, MDIForm1
End Sub

Private Sub Form_Load()
    frmSearch.Height = 1950
    MDIForm1.StatusBar1.Panels(1).Text = "Action: Awaiting Search Request"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: None"
    MDIForm1.StatusBar1.Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.StatusBar1.Panels(1).Text = "Action: None"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: None"
    MDIForm1.StatusBar1.Refresh
End Sub
Sub errRoutine()
    If gstatus <> "1" Or gstatus <> "0" Then
        gstatus = 0
    End If
    Module1.errRTN
End Sub
