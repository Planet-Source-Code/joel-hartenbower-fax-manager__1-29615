VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmUserAdmin 
   Caption         =   "User Administration"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   Icon            =   "frmUserAdmin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFullName 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1800
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1200
      Width           =   735
   End
   Begin VB.CheckBox chkOld 
      Caption         =   "Check1"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   2880
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Access Levels (Future)"
      Height          =   1215
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox txtUsername 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.ListBox lstUser 
      Height          =   3180
      Left            =   3720
      Sorted          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc datUserDB 
      Height          =   330
      Left            =   120
      Top             =   3120
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name:"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmUserAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    Dim varPassword, varUserName, varFullName As String
    
    On Error GoTo errRoutine
    varUserName = EnCryption(UCase(Left(txtUsername.Text, 1)) & LCase(Right(txtUsername.Text, Len(txtUsername.Text) - 1)))
    varPassword = EnCryption(LCase(txtPassword))
    varFullName = EnCryption(txtFullName.Text)
    
    If Len(varPassword) = 0 Then
        MsgBox "User cannot be saved until a password has been specified.", vbInformation, "Password Error"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    gTitle = "Recordset Error"
    gMsg = "A1 - There was a problem opening the recordset."
    gstatus = 0
    
    If chkOld = 0 Then
        gSQL = "SELECT * FROM [User Table]"
    Else
        gSQL = "SELECT * FROM [User Table] WHERE UserName='" & varUserName & "'"
    End If
    datUserDB.RecordSource = gSQL
    datUserDB.Refresh
        
    With datUserDB.Recordset
        If chkOld = 0 Then
            .AddNew
        End If
        .Fields("Username") = varUserName
        .Fields("Password") = varPassword
        .Fields("FullName") = varFullName
        .Update
    End With
    
    txtUsername.Text = ""
    txtPassword.Text = ""
    txtFullName.Text = ""
    chkOld.Value = False
    txtUsername.SetFocus
    SendKeys "{Home}+{End}"
    subRefreshUsers
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
End Sub
Sub subRefreshUsers()
    Dim varPassword, varUserName As String
    Dim varCount, varTotal As Integer
    
    On Error GoTo errRoutine
    gTitle = "Recordset Error"
    gMsg = "A2 - There was a problem opening the recordset."
    gstatus = 0
    gSQL = "SELECT * FROM [User Table]"
    datUserDB.RecordSource = gSQL
    datUserDB.Refresh
    
    varCount = 0
    lstUser.Clear
    datUserDB.Recordset.MoveFirst
    Do While Not datUserDB.Recordset.EOF
        varUserName = EnCryption(datUserDB.Recordset("Username"))
        lstUser.AddItem varUserName
        datUserDB.Recordset.MoveNext
    Loop
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
End Sub

Private Sub Command3_Click()
    txtUsername.Text = ""
    txtPassword.Text = ""
    txtFullName.Text = ""
    txtUsername.SetFocus
    chkOld.Value = False
    SendKeys "{Home}+{End}"
    Command4.Enabled = False
End Sub

Private Sub Command4_Click()
    Dim varPassword, varErrMsg As String
    Dim varUserName As String
    
    If UCase(txtUsername.Text) = "ADMIN" Then
        MsgBox "The Admin account cannot be deleted.  It is used for administrative purposes.", vbInformation, "Cannot Delete"
        Exit Sub
    End If
    
    On Error GoTo errRoutine
    varUserName = UCase(Left(txtUsername.Text, 1)) & LCase(Right(txtUsername.Text, Len(txtUsername.Text) - 1))
    varUserName = EnCryption(varUserName)
    
    gTitle = "Recordset Error"
    gMsg = "A3 - There was a problem opening the recordset."
    gstatus = 0
    gSQL = "SELECT * FROM [User Table] WHERE UserName='" & varUserName & "'"
    datUserDB.RecordSource = gSQL
    datUserDB.Refresh
    
    datUserDB.Recordset.MoveFirst
    If Not datUserDB.Recordset.EOF Then
        datUserDB.Recordset.Delete
    Else
        If Not rdatUserDB.Recordset.EOF Then
            MsgBox "That record was already deleted.", vbInformation, "Unsuccessful"
            Exit Sub
        End If
    End If
    gSQL = "SELECT * FROM [User Table] WHERE UserName='" & varUserName & "'"
    datUserDB.RecordSource = gSQL
    datUserDB.Refresh
    
    datUserDB.Recordset.MoveFirst
    If datUserDB.Recordset.EOF Then
        MsgBox "Record was successfully deleted.", vbOKOnly, "Successful"
    Else
        varErrMsg = Err.Description & Chr(13) & Err.Number
        MsgBox varErrMsg, vbCritical, "Database Error"
        End
    End If
    txtUsername.Text = ""
    txtPassword.Text = ""
    txtFullName.Text = ""
    chkOld.Value = False
    txtUsername.SetFocus
    SendKeys "{Home}+{End}"
    Command4.Enabled = False
    subRefreshUsers
    Exit Sub
    
errRoutine:
    If Err.Number = 3021 Then Resume Next
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
End Sub

Private Sub Form_Load()
    subRefreshUsers
End Sub

Private Sub lstUser_Click()
    Dim varPassword, varUserName As String
    
    On Error Resume Next
    txtUsername.Text = lstUser
    varUserName = lstUser
    varUserName = EnCryption(varUserName)
    
    gTitle = "Recordset Error"
    gMsg = "A4 - There was a problem opening the recordset."
    gstatus = 0
    gSQL = "SELECT * FROM [User Table] WHERE UserName='" & varUserName & "'"
    datUserDB.RecordSource = gSQL
    datUserDB.Refresh
    
    txtPassword.Text = EnCryption(datUserDB.Recordset("Password"))
    txtFullName.Text = EnCryption(datUserDB.Recordset("FullName"))
    chkOld = 1
    Command4.Enabled = True
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
End Sub
Sub errRoutine()
    If gstatus <> "1" Or gstatus <> "0" Then
        gstatus = 0
    End If
    Module1.errRTN
End Sub
