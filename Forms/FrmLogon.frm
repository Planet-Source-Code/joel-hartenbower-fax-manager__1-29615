VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmLogon 
   Caption         =   "Login for Fax Manager"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   Icon            =   "FrmLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc datUserDB 
      Height          =   330
      Left            =   6600
      Top             =   0
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
   Begin VB.CommandButton Command1 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login"
      Height          =   1695
      Left            =   5520
      TabIndex        =   8
      Top             =   3840
      Width           =   3255
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdLogin 
         Caption         =   "&Login"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Username:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Label Label5 
      Caption         =   "AM/PM PC Services, Inc."
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   5280
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Joel W. Hartenbower"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5040
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "Written by:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   3855
   End
   Begin VB.Label lblPlatform 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Windows 95/98/ME/NT/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4575
      TabIndex        =   7
      Top             =   3000
      Width           =   4140
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Version 3.01 (Alpha)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6435
      TabIndex        =   6
      Top             =   3360
      Width           =   2280
   End
   Begin VB.Label lblCompanyProduct 
      AutoSize        =   -1  'True
      Caption         =   "AM/PM PC Serivces, Inc."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2317
      TabIndex        =   5
      Top             =   360
      Width           =   4230
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      Caption         =   "Fax Manager"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   2482
      TabIndex        =   4
      Top             =   795
      Width           =   3960
   End
   Begin VB.Image imgLogo 
      Height          =   2385
      Left            =   240
      Picture         =   "FrmLogon.frx":0442
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "FrmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdLogin_Click()
    Dim varDBName, varUserHold, varPasswordHold, varUserName As String
    Dim varCount, varQuit As Integer
    Dim hKey As Long
    
    On Error GoTo errRoutine
    If Len(txtUsername.Text) = 0 Then
        MsgBox "Please enter a valid Logon and Password.", vbExclamation, AppTitle
        txtUsername.SetFocus
        SendKeys "^+{HOME}"
        Exit Sub
    End If
    
    varUserName = UCase(Left(txtUsername.Text, 1)) & LCase(Right(txtUsername.Text, Len(txtUsername.Text) - 1))
    varUserHold = EnCryption(varUserName)
    varPasswordHold = EnCryption(LCase(txtPassword.Text))
    
    gTitle = "Recordset Error"
    gMsg = "L1 - There was a problem opening the recordset."
    gstatus = 0
    
    gSQL = "SELECT * FROM [User Table] WHERE Username = '" & varUserHold & "'"
    datUserDB.RecordSource = gSQL
    datUserDB.Refresh
    
    If datUserDB.Recordset.EOF Then
        MsgBox "Please enter a valid Logon and Password.", vbExclamation, AppTitle
        txtUsername.SetFocus
        SendKeys "^+{HOME}"
        Exit Sub
    Else
        If EnCryption(LCase(txtPassword.Text)) <> datUserDB.Recordset("Password") Then
            MsgBox "Your password does not match, please try again.", vbInformation, "Password Mismatch"
            txtPassword.SetFocus
            SendKeys "^+{HOME}"
            Exit Sub
        End If
        If LCase(varUserName) = LCase(txtPassword.Text) Then
            MsgBox "You must change your password.  Choose Miscellaneous from the menu bar then click on change password.", vbInformation, "Change Password"
        End If
        
        gUserName = varUserName
        hKey = &H80000001
        gDatabaseName = GetRegValue(hKey, "Software\ODBC\ODBC.INI\FaxManagerDB", "DBQ", "Not Found")
        If gDatabaseName <> "Not Found" Then
            varCount = 0
            Do While Left(varDBNameHold, 1) <> "\" And varCount < Len(gDatabaseName)
                varDBNameHold = Right(gDatabaseName, 34 - varCount)
                varCount = varCount + 1
            Loop
            varDBName = Left(varDBNameHold, Len(varDBNameHold) - 4)
            varDBName = Right(varDBName, Len(varDBName) - 1)
            MDIForm1.StatusBar1.Panels(3).Text = "Open DB: " & varDBName
            MDIForm1.StatusBar1.Refresh
            EnableItems
        Else
            MDIForm1.StatusBar1.Panels(3).Text = "Open DB: -- No Database Selected --"
            MDIForm1.StatusBar1.Refresh
            DisableItems
        End If
        MDIForm1.StatusBar1.Panels(5).Text = "Username: " & gUserName
        MDIForm1.StatusBar1.Refresh
        MDIForm1.Show
    End If
    Unload Me
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
End Sub

Private Sub Command1_Click()
    End
End Sub

Private Sub Form_Load()
    Module1.app_path
    gUserDatabase = gApp_Path & "Fax Manager.mdb"
    gTitle = "Registery Error"
    gMsg = "Could not register the database in the system."
    gstatus = 1
    strAttributes = "DBQ=" & gUserDatabase
    DBEngine.RegisterDatabase "FaxManagerUser", "Microsoft Access Driver (*.mdb)", True, strAttributes
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If gUserName = "" Then
        End
    Else
        If gUserName = "Admin" Or gUserName = "Joel" Or LCase(gUserName) = "jrgrade" Then
            MDIForm1.mnuAdmin.Visible = True
            MDIForm1.mnuAdmin.Enabled = True
        Else
            MDIForm1.mnuAdmin.Visible = False
        End If
    End If
End Sub
Sub errRoutine()
    If gstatus <> "1" Or gstatus <> "0" Then
        gstatus = 0
    End If
    Module1.errRTN
End Sub
Sub EnableItems()
    MDIForm1.mnuFileItem(1).Enabled = True
    MDIForm1.mnuFileItem(0).Enabled = False
    MDIForm1.mnuMaintenece.Enabled = True
    MDIForm1.mnuReports.Enabled = True
    MDIForm1.mnuMisc.Enabled = True
    MDIForm1.mnuMiscItem(0).Enabled = True
    MDIForm1.mnuAdmin.Enabled = True
    MDIForm1.Toolbar1.Buttons(1).Enabled = False
    MDIForm1.Toolbar1.Buttons(2).Enabled = True
    MDIForm1.Toolbar1.Buttons(4).Enabled = True
    MDIForm1.Toolbar1.Buttons(5).Enabled = True
    MDIForm1.Toolbar1.Buttons(6).Enabled = True
    MDIForm1.Toolbar1.Buttons(7).Enabled = True
End Sub
Sub DisableItems()
    MDIForm1.mnuFileItem(1).Enabled = False
    MDIForm1.mnuFileItem(0).Enabled = True
    MDIForm1.mnuMaintenece.Enabled = False
    MDIForm1.mnuReports.Enabled = False
    MDIForm1.mnuMisc.Enabled = False
    MDIForm1.mnuAdmin.Enabled = False
    MDIForm1.Toolbar1.Buttons(1).Enabled = True
    MDIForm1.Toolbar1.Buttons(2).Enabled = False
    MDIForm1.Toolbar1.Buttons(4).Enabled = False
    MDIForm1.Toolbar1.Buttons(5).Enabled = False
    MDIForm1.Toolbar1.Buttons(6).Enabled = False
    MDIForm1.Toolbar1.Buttons(7).Enabled = False
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdLogin_Click
End Sub
Private Sub txtUsername_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtPassword.SetFocus
End Sub
