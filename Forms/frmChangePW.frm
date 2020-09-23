VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmChangePW 
   Caption         =   "Change Password"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmChangePW.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Change"
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtConfirm 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtNewPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox txtOldPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin MSAdodcLib.Adodc datUserDB 
      Height          =   330
      Left            =   0
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
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4440
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirm Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "New Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Old Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmChangePW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim SQL, varOldPWHold, varNewPWHold As String
    
    On Error Resume Next
    varNewPWHold = txtNewPW.Text

    If varNewPWHold = txtConfirm.Text Then
        gSQL = "SELECT * FROM [User Table] WHERE UserName='" & EnCryption(gUserName) & "'"
        datUserDB.RecordSource = gSQL
        datUserDB.Refresh
        
        If Err.Number <> 0 Then
            Module1.errRTN "Password Error", "There was a problem opening the Recordset.", Err.Number, Err.Description, 0
            Exit Sub
        End If
    
        If EnCryption(LCase(txtOldPW.Text)) = datUserDB.Recordset("Password") Then
            With datUserDB.Recordset
                .Fields("Password") = EnCryption(LCase(txtNewPW.Text))
                .Update
            End With
            If Err.Number = 0 Then
                Unload Me
            End If
            If Err.Number <> 0 Then
                MsgBox "There was a problem saving the password.", vbCritical, "Password Error"
                Exit Sub
            End If
        Else
            MsgBox "Your old password does not match what is on file, please try again.", vbInformation, "Password Error"
            txtOldPW.SetFocus
            SendKeys "^+{HOME}"
            Exit Sub
        End If
    Else
        MsgBox "Your new password and the confirmation do not match, please try again.", vbInformation, "Password Error"
        txtNewPW.Text = ""
        txtConfirm.Text = ""
        txtNewPW.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
