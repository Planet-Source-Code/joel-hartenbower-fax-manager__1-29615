VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmWinFax 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export (Access)"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtDBName 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Choose Export Database"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CheckBox ckExport 
      Caption         =   "Export All Records"
      Height          =   195
      Left            =   2002
      TabIndex        =   3
      Top             =   1080
      Width           =   1620
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start Export"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin MSMask.MaskEdBox txtNumberRecords 
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtLastRecord 
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Finished"
      Height          =   375
      Left            =   1785
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtDBPath 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   720
      Width           =   4095
   End
   Begin MSAdodcLib.Adodc datExportDB 
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
      Connect         =   "DSN=FaxManagerExportDB"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "FaxManagerExportDB"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "datExportDB"
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
   Begin MSAdodcLib.Adodc datFaxDB 
      Height          =   330
      Left            =   0
      Top             =   360
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
   Begin MSAdodcLib.Adodc datUserDB 
      Height          =   330
      Left            =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Connect         =   "DSN=FaxManagerUser"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "FaxManagerUser"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "datUserDB"
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
   Begin VB.Label lblDelete 
      Alignment       =   2  'Center
      Caption         =   "Deleting Records"
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblWait 
      Alignment       =   2  'Center
      Caption         =   "Setting up the databases, Please Wait..."
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Database Path:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Database Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lblRecord 
      Caption         =   "1"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Processing Record:"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblPercent 
      Alignment       =   1  'Right Justify
      Caption         =   "0%"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "End of last export:"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Number of Records to Export:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmWinFax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varUserName As String

Private Sub Command1_Click()
    Dim dblTotalRecords, dblCount As Double
    
    Me.Height = 2655
    Command1.Visible = False
    Command3.Visible = False
    lblRecord.Caption = "0"
    lblPercent.Caption = "0%"
    
    On Error GoTo errRoutine
    txtNumberRecords.Enabled = False
    txtLastRecord.Enabled = False
    ckExport.Enabled = False
    ProgressBar1.Visible = False
    lblPercent.Visible = False
    Label3.Visible = False
    lblRecord.Visible = False
    lblWait.Visible = True
    lblWait.Refresh
    
    varUserName = EnCryption(gUserName)
    gSQL = "SELECT * FROM [User Table] WHERE Username='" & varUserName & "'"
    datUserDB.RecordSource = gSQL
    datUserDB.Refresh
    
    gSQL = "SELECT * FROM [Phone Book Table]"
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh
    
    gSQL = "SELECT * FROM [ZetaFax]"
    datExportDB.RecordSource = gSQL
    datExportDB.Refresh
    
    lblDelete.Visible = True
    lblDelete.Refresh
    ProgressBar1.Visible = True
    lblPercent.Visible = True
    dblCount = 0
    dblTotalRecords = datExportDB.Recordset.RecordCount
    If datExportDB.Recordset.RecordCount > 0 Then
        Do While datExportDB.Recordset.EOF = False
            dblCount = dblCount + 1
            ProgressBar1.Value = (dblCount / dblTotalRecords) * 100
            lblPercent.Caption = Round(((dblCount / dblTotalRecords) * 100), 0) & "%"
            lblPercent.Refresh
            datExportDB.Recordset.Delete
            datExportDB.Recordset.MoveNext
        Loop
    End If
    
    If ckExport.Value <> 0 Then
        dblTotalRecords = datFaxDB.Recordset.RecordCount
    Else
        dblTotalRecords = txtNumberRecords.Text
        datFaxDB.Recordset.Move (Int(txtLastRecord.Text) + 1)
    End If
    
    dblCount = 0
    ProgressBar1.Value = 0
    lblPercent.Caption = "0%"
    Label3.Visible = True
    lblRecord.Visible = True
    lblWait.Caption = "Exporting Records, Please Wait..."
    lblDelete.Visible = False
    
    Do While datFaxDB.Recordset.EOF = False
        If dblCount = Int(txtNumberRecords) And ckExport.Value = 0 Then
            Exit Do
        End If
        dblCount = dblCount + 1
        ProgressBar1.Value = (dblCount / dblTotalRecords) * 100
        lblPercent.Caption = Round(((dblCount / dblTotalRecords) * 100), 0) & "%"
        If ckExport.Value <> 0 Then
            lblRecord.Caption = dblCount + Int(txtLastRecord.Text)
        Else
            lblRecord.Caption = dblCount
        End If
        lblRecord.Refresh
        lblPercent.Refresh
        With datExportDB.Recordset
            .AddNew
            .Fields("FullName") = datFaxDB.Recordset("FirstName") & " " & datFaxDB.Recordset("LastName")
            .Fields("FirstName") = datFaxDB.Recordset("FirstName")
            .Fields("LastName") = datFaxDB.Recordset("LastName")
            .Fields("Company") = datFaxDB.Recordset("Company")
            .Fields("FaxNumber") = datFaxDB.Recordset("FaxNumber")
            .Update
        End With
        datFaxDB.Recordset("UsedCount") = datFaxDB.Recordset("UsedCount") + 1
        datFaxDB.Recordset("LastUsed") = Date
        datFaxDB.Recordset.MoveNext
        lblRecord.Caption = dblCount + Int(txtLastRecord.Text)
    Loop
    ProgressBar1.Value = 100
    lblPercent.Caption = "100%"
    ProgressBar1.Refresh
    lblPercent.Refresh
    
    If ckExport.Value = 0 And dblCount + Int(txtLastRecord.Text) < datFaxDB.Recordset.RecordCount Then
        datUserDB.Recordset("LastExportRecord") = dblCount + Int(txtLastRecord.Text)
        datUserDB.Recordset.Update
        txtLastRecord.Text = dblCount + Int(txtLastRecord.Text)
    Else
        datUserDB.Recordset("LastExportRecord") = 0
        datUserDB.Recordset.Update
    End If
    lblWait.Visible = False
    Command2.Visible = True
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
    Unload Me
End Sub
Sub errRoutine()
    If gstatus <> "1" Or gstatus <> "0" Then
        gstatus = 0
    End If
    Module1.errRTN
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim intCount As Integer
    Dim varDBName As String
    
    CommonDialog1.CancelError = True
    On Error GoTo errRoutine
    
    CommonDialog1.Filter = "Database (*.mdb) |*.mdb| All Files (*.*) |*.*"
    Me.MousePointer = 11
    CommonDialog1.Action = 1
    Me.MousePointer = 0
    gExportDB = CommonDialog1.FileName
    If Err.Number <> 3059 Then
        Dim dbsRegister As Database
        Dim strAttributes As String
        
        'Register the new database for ODBC
        gTitle = "Registery Error"
        gMsg = "Could not register the database in the system."
        strAttributes = "DBQ=" & gExportDB
        DBEngine.RegisterDatabase "FaxManagerExportDB", "Microsoft Access Driver (*.mdb)", True, strAttributes
        
        'Strip Name
        intCount = 0
        Do While intCount < Len(gExportDB) + 1 And Left(varDBName, 1) <> "\"
            intCount = intCount + 1
            varDBName = Right(gExportDB, intCount)
        Loop
        txtDBName.Text = Right(varDBName, intCount - 1)
        txtDBPath.Text = Left(gExportDB, (Len(gExportDB) - intCount) + 1)
    End If
    Exit Sub

errRoutine:
    If Err.Number = 3059 Then Resume Next
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
End Sub

Private Sub Form_Load()
    Dim intCount As Integer
    Dim hKey As Long
    Dim varDBName As String
    
    txtNumberRecords.Text = "1000"
    txtLastRecord.Text = "0"
    
    Me.Height = 2130
    Me.Width = 5715
    
    On Error Resume Next
    hKey = &H80000001
    gExportDB = GetRegValue(hKey, "Software\ODBC\ODBC.INI\FaxManagerExportDB", "DBQ", "Not Found")
    
    If gExportDB <> "Not Found" Then
        'Strip Name
        intCount = 0
        Do While intCount < Len(gExportDB) + 1 And Left(varDBName, 1) <> "\"
            intCount = intCount + 1
            varDBName = Right(gExportDB, intCount)
        Loop
        txtDBName.Text = Right(varDBName, intCount - 1)
        txtDBPath.Text = Left(gExportDB, (Len(gExportDB) - intCount) + 1)
    Else
        txtDBName.Text = ""
        txtDBPath.Text = ""
    End If
    
    varUserName = EnCryption(gUserName)
    gSQL = "SELECT * FROM [User Table] WHERE UserName='" & varUserName & "'"
    datUserDB.RecordSource = gSQL
    datUserDB.Refresh
    
    If datUserDB.Recordset.EOF = False Then
        If datUserDB.Recordset("LastExportRecord") = Null Then
            txtLastRecord.Text = 0
        Else
            txtLastRecord.Text = datUserDB.Recordset("LastExportRecord")
        End If
    End If
    txtNumberRecords.SetFocus
    SendKeys "^+{END}"
End Sub
