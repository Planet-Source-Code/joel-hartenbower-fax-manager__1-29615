VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmQuickDelVerify 
   Caption         =   "Delete Verification"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   Icon            =   "frmQuckDelVerify.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCurrent 
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   3360
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdConfirm 
      Caption         =   "&Confirm and Continue"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdSkip 
      Caption         =   "&Skip"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Re-enter Fax Number"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtCompany 
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   3255
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
   Begin MSAdodcLib.Adodc datFaxDel 
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
      Caption         =   "datFaxDel"
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
   Begin VB.Label lblTotal 
      Caption         =   "0"
      Height          =   255
      Left            =   2700
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Total Matches:"
      Height          =   255
      Left            =   1500
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Fax Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Full Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmQuickDelVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirm_Click()
    Dim strSearchText As String
    Dim Msg As String
    Dim Result, Looping As Integer
    
    On Error GoTo errRoutine
    Looping = 0
    If lblTotal > 0 Then Looping = 1
    
    gTitle = "Recordset Error"
    gMsg = "QD1 - There was a problem opening the recordset."
    gstatus = 0
    gSQL = "SELECT * FROM [Phone Book Table] WHERE FaxNumber = '" & txtPhone.Text & "'"
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh
    
    If datFaxDB.Recordset.EOF = False Then
        datFaxDB.Recordset.MoveFirst
        gSQL = "SELECT * FROM [Deleted Phone Book Table]"
        datFaxDel.RecordSource = gSQL
        datFaxDel.Refresh
        gTitle = "Recordset Error"
        gMsg = "QD2 - There was a problem moving the record to the deleted backup."
        gstatus = 0
        With datFaxDel.Recordset
            .AddNew
            .Fields("Company") = datFaxDB.Recordset("Company")
            .Fields("LastName") = datFaxDB.Recordset("LastName")
            .Fields("FirstName") = datFaxDB.Recordset("FirstName")
            .Fields("Title") = datFaxDB.Recordset("Title")
            .Fields("Address1") = datFaxDB.Recordset("Address1")
            .Fields("Address2") = datFaxDB.Recordset("Address2")
            .Fields("City") = datFaxDB.Recordset("City")
            .Fields("State") = datFaxDB.Recordset("State")
            .Fields("Zip") = datFaxDB.Recordset("Zip")
            .Fields("FaxNumber") = datFaxDB.Recordset("FaxNumber")
            .Fields("PhoneNumber") = datFaxDB.Recordset("PhoneNumber")
            .Fields("Email") = datFaxDB.Recordset("Email")
            .Fields("Category1") = datFaxDB.Recordset("Category1")
            .Fields("Category2") = datFaxDB.Recordset("Category2")
            .Fields("DateDel") = Date
            .Fields("DelBy") = gUserName
            .Update
        End With
        gTitle = "Recordset Error"
        gMsg = "QD3 - There was a problem deleting the current record."
        gstatus = 0
        datFaxDB.Recordset.Delete
        MsgBox "Delete was successful.", vbInformation, "Successful"
    Else
        MsgBox "That record was not found or was previously deleted.", vbInformation, "Record Not Found"
        Exit Sub
    End If
    If Looping = 0 Then
        Unload Me
    Else
        lblTotal.Caption = lblTotal.Caption - 1
        lblTotal.Refresh
        If lblTotal.Caption = 0 Then
            Unload Me
        Else
            gTitle = "Recordset Error"
            gMsg = "QD4 - There was a problem finding the next record."
            gstatus = 0
            datFaxDB.Recordset.MoveNext
            txtCompany.Text = datFaxDB.Recordset("Company")
            txtName.Text = datFaxDB.Recordset("FirstName") & " " & datFaxDB.Recordset("LastName")
            varPhone = Right(datFaxDB.Recordset("FaxNumber"), 10)
            txtPhone.Text = datFaxDB.Recordset("FaxNumber")
            varPhoneHold = "(" & Left(varPhone, 3) & ") "
            varPhone = Right(varPhone, 7)
            txtNumber.Text = varPhoneHold & Left(varPhone, 3) & "-" & Right(varPhone, 4)
            If lblTotal.Caption > 1 Then
                cmdSkip.Visible = True
                txtCurrent.Text = 1
            Else
                cmdSkip.Visible = False
                cmdConfirm.Caption = "&Confirm Delete"
            End If
        End If
    End If
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
    Unload Me
End Sub

Private Sub cmdSkip_Click()
    Dim strSearchText As String
    Dim Msg As String
    Dim Result, Looping As Integer
    
    On Error GoTo errRoutine
    Looping = 0
    If lblTotal > 0 Then Looping = 1
    
    gTitle = "Recordset Error"
    gMsg = "QD6 - There was a problem opening the recordset."
    gstatus = 0
    gSQL = "SELECT * FROM [Phone Book Table] WHERE FaxNumber = '" & txtPhone.Text & "'"
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh

    If datFaxDB.Recordset.EOF = False Then
        Looping = 1
        If txtCurrent.Text = lblTotal.Caption Then
            txtCurrent.Text = 1
        Else
            txtCurrent.Text = txtCurrent.Text + 1
        End If
        datFaxDB.Recordset.MoveFirst
        Do While Looping < txtCurrent.Text
            datFaxDB.Recordset.MoveNext
            Looping = Looping + 1
        Loop
        If lblTotal.Caption = 0 Then
            Unload Me
        Else
            txtCompany.Text = datFaxDB.Recordset("Company")
            txtName.Text = datFaxDB.Recordset("FirstName") & " " & datFaxDB.Recordset("LastName")
            varPhone = Right(datFaxDB.Recordset("FaxNumber"), 10)
            txtPhone.Text = datFaxDB.Recordset("FaxNumber")
            varPhoneHold = "(" & Left(varPhone, 3) & ") "
            varPhone = Right(varPhone, 7)
            txtNumber.Text = varPhoneHold & Left(varPhone, 3) & "-" & Right(varPhone, 4)
            If lblTotal.Caption > 1 Then
                cmdSkip.Visible = True
            Else
                cmdConfirm.Caption = "&Confirm Delete"
            End If
            Me.Refresh
        End If
    Else
        MsgBox "No more records", vbOKOnly, "Record Not Found"
        Exit Sub
    End If
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
    Unload Me
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim SQL, varPhone, varPhoneHold As String
    
    On Error GoTo errRoutine
    gTitle = "Recordset Error"
    gMsg = "QD7 - There was a problem opening the recordset."
    gstatus = 0
    gSQL = "SELECT * FROM [Phone Book Table] WHERE FaxNumber = '" & frmQuickDelSearch.txtSearch.Text & "'"
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh
    
    If datFaxDB.Recordset.RecordCount = 0 Then
        MsgBox "No records found", vbInformation, "No Record"
        Unload frmQuickDelVerify
        Exit Sub
    End If
    If datFaxDB.Recordset.RecordCount > 0 Then lblTotal.Caption = datFaxDB.Recordset.RecordCount
    
    If datFaxDB.Recordset.RecordCount > 1 Then
        cmdSkip.Visible = True
        txtCurrent.Text = 1
    Else
        cmdConfirm.Caption = "&Confirm Delete"
    End If
    
    lblTotal.Refresh
    datFaxDB.Recordset.MoveFirst
    txtCompany.Text = datFaxDB.Recordset("Company")
    txtName.Text = datFaxDB.Recordset("FirstName") & " " & datFaxDB.Recordset("LastName")
    varPhone = Right(datFaxDB.Recordset("FaxNumber"), 10)
    txtPhone.Text = datFaxDB.Recordset("FaxNumber")
    varPhoneHold = "(" & Left(varPhone, 3) & ") "
    varPhone = Right(varPhone, 7)
    txtNumber.Text = varPhoneHold & Left(varPhone, 3) & "-" & Right(varPhone, 4)
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmQuickDelSearch.txtSearch.SetFocus
    SendKeys "^+{HOME}"
End Sub
Sub errRoutine()
    If gstatus <> "1" Or gstatus <> "0" Then
        gstatus = 0
    End If
    Module1.errRTN
End Sub
