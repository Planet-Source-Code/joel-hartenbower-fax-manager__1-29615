VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAddEdit 
   Caption         =   "Data Entry/Edit"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   Icon            =   "frmAddEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   3480
      TabIndex        =   30
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtAddUser 
      Height          =   285
      Left            =   0
      TabIndex        =   29
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtAddDate 
      Height          =   285
      Left            =   0
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   27
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   5880
      TabIndex        =   26
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtDateMod 
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   2760
      Width           =   5775
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   9
      Top             =   1800
      Width           =   5775
   End
   Begin VB.TextBox txtRecordID 
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtCompany 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   4320
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   5775
   End
   Begin VB.TextBox txtAddress1 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   5775
   End
   Begin VB.TextBox txtAddress2 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   5775
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox txtState 
      Height          =   285
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtZip 
      Height          =   285
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   8
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtFax 
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   10
      Top             =   2040
      Width           =   5775
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Top             =   2280
      Width           =   5775
   End
   Begin VB.TextBox Text13 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1440
      TabIndex        =   13
      Top             =   3000
      Width           =   5775
   End
   Begin MSAdodcLib.Adodc datFaxDB 
      Height          =   330
      Left            =   0
      Top             =   3360
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
   Begin VB.Label lblWorking 
      Alignment       =   2  'Center
      Caption         =   "Working, Please Wait..."
      Height          =   255
      Left            =   0
      TabIndex        =   31
      Top             =   3600
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Variable 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Fax Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Voice Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Title:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "City/State/Zip:"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "First/Last Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Company:"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Email Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Variable 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo errRoutine
    
    lblWorking.Visible = True
    gTitle = "Recordset Error"
    gMsg = "AE2 - There was a problem opening the recordset."
    gstatus = 0
    gSQL = "SELECT * FROM [Deleted Phone Book Table]"
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh
    
    gTitle = "Recordset Error"
    gMsg = "AE2 - There was a problem moving the current record to the backup deleted table."
    gstatus = 1
    With datFaxDB.Recordset
        .AddNew
        .Fields("Company") = Me.txtCompany
        If Me.txtTitle <> "" Then
            .Fields("Title") = Me.txtTitle
        End If
        If Me.txtFirstName <> "" Then
            .Fields("firstname") = Me.txtFirstName
        End If
        If Me.txtLastName <> "" Then
            .Fields("LastName") = Me.txtLastName
        End If
        If Me.txtAddress1 <> "" Then
            .Fields("Address1") = Me.txtAddress1
        End If
        If Me.txtAddress2 <> "" Then
            .Fields("Address2") = Me.txtAddress2
        End If
        If Me.txtCity <> "" Then
            .Fields("City") = Me.txtCity
        End If
        If Me.txtState <> "" Then
            .Fields("State") = Me.txtState
        End If
        If Me.txtZip <> "" Then
            .Fields("Zip") = Me.txtZip
        End If
        If Me.txtPhone <> "" Then
            .Fields("PhoneNumber") = Me.txtPhone
        End If
        If Me.txtFax <> "" Then
            .Fields("FaxNumber") = Me.txtFax
        End If
        If Me.txtEmail <> "" Then
            .Fields("Email") = Me.txtEmail
        End If
        .Fields("DateDel") = Date
        .Fields("DelBy") = gUserName
        .Update
    End With
    gTitle = "Recordset Error"
    gMsg = "AE2 - There was a problem deleting the current record."
    gstatus = 1
    gSQL = "SELECT * FROM [Phone Book Table] WHERE PID=" & Me.txtRecordID
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh
    datFaxDB.Recordset.Delete
    lblWorking.Visible = False
    MsgBox "Record was successfully deleted.", vbInformation, "Successful"
    Unload Me
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
    lblWorking.Visible = False
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo errRoutine
    
    lblWorking.Visible = True
    If Len(txtFax) = 0 Or Len(txtFax) < 10 Then
        MsgBox "You must enter a fax number before saving.  The phone number must consist of 10 digits.", vbOKOnly, "Missing Data"
        Exit Sub
    End If
    If Len(txtFax) > 10 Or Len(txtPhone) > 10 Then
        MsgBox "Phone numbers must be entered without any punctuation (ie. 9133852676).", vbOKOnly, "Format Error"
        Exit Sub
    End If
    
    gstatus = 0
    If gRecordID > 0 Then
        gstatus = 1
    End If
    If gstatus = 1 Then
        gTitle = "Recordset Error"
        gMsg = "AE3 - There was a problem opening the recordset."
        gstatus = 0
        gSQL = "SELECT * FROM [Phone Book Table] WHERE PID = " & gRecordID
        datFaxDB.RecordSource = gSQL
        datFaxDB.Refresh
    Else
        gTitle = "Recordset Error"
        gMsg = "AE4 - There was a problem opening the recordset."
        gstatus = 0
        gSQL = "SELECT * FROM [Phone Book Table]"
        datFaxDB.RecordSource = gSQL
        datFaxDB.Refresh
    End If
    
    gTitle = "Recordset Error"
    gMsg = "AE5 - There was a problem saving the current record."
    gstatus = 0
    With datFaxDB.Recordset
        If Len(txtRecordID.Text) = 0 Then
            .AddNew
        End If
        .Fields("Company") = Me.txtCompany
        If Me.txtTitle <> "" Then
            .Fields("Title") = Me.txtTitle
        End If
        If Me.txtFirstName <> "" Then
            .Fields("firstname") = Me.txtFirstName
        End If
        If Me.txtLastName <> "" Then
            .Fields("LastName") = Me.txtLastName
        End If
        If Me.txtAddress1 <> "" Then
            .Fields("Address1") = Me.txtAddress1
        End If
        If Me.txtAddress2 <> "" Then
            .Fields("Address2") = Me.txtAddress2
        End If
        If Me.txtCity <> "" Then
            .Fields("City") = Me.txtCity
        End If
        If Me.txtState <> "" Then
            .Fields("State") = Me.txtState
        End If
        If Me.txtZip <> "" Then
            .Fields("Zip") = Me.txtZip
        End If
        If Me.txtPhone <> "" Then
            .Fields("PhoneNumber") = Me.txtPhone
        End If
        If Me.txtFax <> "" Then
            .Fields("FaxNumber") = Me.txtFax
        End If
        If Me.txtEmail <> "" Then
            .Fields("Email") = Me.txtEmail
        End If
        If txtRecordID.Text = 0 Then
            .Fields("AddedBy") = Me.txtAddUser
            .Fields("DateAdded") = Me.txtAddDate
        Else
            .Fields("DateMod") = Me.txtDateMod
            .Fields("ModBy") = Me.txtAddUser
        End If
        .Update
    End With
    lblWorking.Visible = False
    MsgBox "Record was saved successfully.", vbOKOnly, "Successful"
    Unload Me
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
    lblWorking.Visible = False
End Sub

Private Sub Form_Load()
    On Error GoTo errRoutine
    
    If gRecordID > 0 Then
        Me.txtRecordID = gRecordID
    Else
        Me.txtRecordID = 0
    End If
    
    Me.txtAddDate = Date
    Me.txtAddUser = gUserName
    Unload frmSearch
    If gRecordID > 0 Then
        gTitle = "Recordset Error"
        gMsg = "AE1 - There was a problem opening the recordset."
        gstatus = 0
        gSQL = "SELECT * FROM [Phone Book Table] WHERE PID = " & gRecordID
        datFaxDB.RecordSource = gSQL
        datFaxDB.Refresh
        
        If datFaxDB.Recordset.EOF = False Then
            With datFaxDB
                Me.txtCompany.Text = .Recordset("Company")
                If .Recordset("Title") <> "" Then
                    Me.txtTitle.Text = .Recordset("Title")
                End If
                Me.txtFirstName.Text = .Recordset("FirstName")
                Me.txtLastName.Text = .Recordset("LastName")
                If .Recordset("Address1") <> "" Then
                    Me.txtAddress1.Text = .Recordset("Address1")
                End If
                If .Recordset("Address2") <> "" Then
                    Me.txtAddress2.Text = .Recordset("Address2")
                End If
                If .Recordset("City") <> "" Then
                    Me.txtCity.Text = .Recordset("City")
                End If
                If .Recordset("State") <> "" Then
                    Me.txtState.Text = .Recordset("State")
                End If
                If .Recordset("Zip") <> "" Then
                    Me.txtZip.Text = .Recordset("Zip")
                End If
                If .Recordset("PhoneNumber") <> "" Then
                    Me.txtPhone.Text = .Recordset("PhoneNumber")
                End If
                If .Recordset("FaxNumber") <> "" Then
                    Me.txtFax.Text = .Recordset("FaxNumber")
                End If
                If .Recordset("Email") <> "" Then
                    Me.txtEmail.Text = .Recordset("Email")
                End If
                Me.txtDateMod.Text = Date
            End With
        End If
    End If
    If Me.txtRecordID > 0 Then
        cmdSave.Caption = "&Update"
        cmdDelete.Visible = True
    Else
        cmdDelete.Visible = False
    End If
    MDIForm1.StatusBar1.Panels(1).Text = "Action: Awaiting Data Entry/Edit/Delete"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: None"
    MDIForm1.StatusBar1.Refresh
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
    Unload Me
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
Private Sub txtFax_LostFocus()
    If Len(txtFax.Text) > 0 Then
        If IsNumeric(Me.txtFax.Text) = False Then
            MsgBox "The phone number must be in the format of 9136632451"
            Me.txtFax.SetFocus
            SendKeys "^+{HOME}"
        End If
    End If
End Sub

Private Sub txtPhone_LostFocus()
    If Len(txtPhone.Text) > 0 Then
        If IsNumeric(Me.txtPhone.Text) = False Then
            MsgBox "The phone number must be in the format of 9136632451"
            Me.txtPhone.SetFocus
            SendKeys "^+{HOME}"
        End If
    End If
End Sub
