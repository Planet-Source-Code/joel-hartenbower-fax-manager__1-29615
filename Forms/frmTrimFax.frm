VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTrimFax 
   Caption         =   "Fax number triming"
   ClientHeight    =   2235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5085
   Icon            =   "frmTrimFax.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&No"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Yes"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Working..."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "DO YOU WISH TO CONTINUE?"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "This will trim the *67, 1 or any other unneeded items from the fax number.  This proceedure will be unreversable. "
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmTrimFax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim dbs As Database
    Dim rstFaxTrim As Recordset
    Dim intTotalRecords As Integer
    Dim intCount, intCountHold As Integer
    
    Me.Height = 2640
    Me.Command1.Enabled = False
    Me.Command2.Enabled = False
    Me.Refresh
    
    Set dbs = OpenDatabase(gDatabaseName)
    On Error Resume Next
    lblDescription.Caption = "Pre-calculations be performened"
    SQL = "Phone Book Table"
    Set rstFaxTrim = dbs.OpenRecordset(SQL)
    If Err.Number <> 0 Then
        Module1.ErrRTN "Dataset Error", "There was a problem opening the Recordset.", Err.Number, Err.Description, 0
        Exit Sub
    End If
    intTotalRecords = rstFaxTrim.RecordCount
    rstFaxTrim.Close
    SQL = "Deleted Phone Book Table"
    Set rstFaxTrim = dbs.OpenRecordset(SQL)
    If Err.Number <> 0 Then
        Module1.ErrRTN "Dataset Error", "There was a problem opening the Recordset.", Err.Number, Err.Description, 0
        Exit Sub
    End If
    intTotalRecords = intTotalRecords + rstFaxTrim.RecordCount
    MDIForm1.StatusBar1.Panels(1).Text = "Action: Processing Trim Request"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: 0 of 0"
    MDIForm1.StatusBar1.Refresh
    intCount = 0
    SQL = "SELECT * FROM [Phone Book Table]"
    Set rstFaxTrim = dbs.OpenRecordset(SQL)
    If Err.Number <> 0 Then
        Module1.ErrRTN "Dataset Error", "There was a problem opening the Recordset.", Err.Number, Err.Description, 0
        Exit Sub
    End If
    rstFaxTrim.MoveFirst
    Do While Not rstFaxTrim.EOF And intCount < intTotalRecords
        intCount = intCount + 1
        ProgressBar1.Value = (intCount / intTotalRecords) * 100
        lblDescription.Caption = "Record Processed: " & intCount & " of " & intTotalRecords
        If intCount = intCountHold + 100 Then
            MDIForm1.StatusBar1.Panels(2).Text = "Status: " & intCount & " of " & intTotalRecords
            MDIForm1.StatusBar1.Refresh
            intCountHold = intCount
        End If
        lblDescription.Refresh
        
        With rstFaxTrim
            .Edit
            .Fields("FaxNumber") = Right(rstFaxTrim.Fields("FaxNumber"), 10)
            .Update
        End With
        rstFaxTrim.MoveNext
    Loop
    rstFaxTrim.Close
    SQL = "SELECT * FROM [Deleted Phone Book Table]"
    Set rstFaxTrim = dbs.OpenRecordset(SQL)
    If Err.Number <> 0 Then
        Module1.ErrRTN "Dataset Error", "There was a problem opening the Recordset.", Err.Number, Err.Description, 0
        Exit Sub
    End If
    rstFaxTrim.MoveFirst
    Do While Not rstFaxTrim.EOF And intCount < intTotalRecords
        intCount = intCount + 1
        ProgressBar1.Value = (intCount / intTotalRecords) * 100
        lblDescription.Caption = "Record Processed: " & intCount & " of " & intTotalRecords
        If intCount = intCountHold + 100 Then
            MDIForm1.StatusBar1.Panels(2).Text = "Status: " & intCount & " of " & intTotalRecords
            MDIForm1.StatusBar1.Refresh
            intCountHold = intCount
        End If
        lblDescription.Refresh
        With rstFaxTrim
            .Edit
            .Fields("FaxNumber") = Right(rstFaxTrim.Fields("FaxNumber"), 10)
            .Update
        End With
        rstFaxTrim.MoveNext
    Loop
    rstFaxTrim.Close
    MsgBox "Changes were successful.", vbOKOnly, "Successful"
    If Err.Number <> 0 Then
        Module1.ErrRTN "Dataset Error", "There was a problem opening the Recordset.", Err.Number, Err.Description, 0
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Height = 1995
    MDIForm1.StatusBar1.Panels(1).Text = "Action: Awaiting confirmation for trim"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: None"
    MDIForm1.StatusBar1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.StatusBar1.Panels(1).Text = "Action: None"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: None"
    MDIForm1.StatusBar1.Refresh
End Sub
