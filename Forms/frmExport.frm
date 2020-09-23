VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExport 
   Caption         =   "Export Records (Text)"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Area4 
      Height          =   285
      Left            =   2880
      TabIndex        =   13
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Area2 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Text            =   "913"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Area3 
      Height          =   285
      Left            =   2400
      TabIndex        =   12
      Text            =   "816"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Area1 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Text            =   "785"
      Top             =   480
      Width           =   375
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Export"
      Height          =   255
      Left            =   6000
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtDirectory 
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   6495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Browse"
      Height          =   255
      Left            =   3480
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6840
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Export As..."
      FileName        =   "PhoneList.txt"
      InitDir         =   "T:\Fax List"
   End
   Begin VB.Label Label3 
      Caption         =   "Local Area Codes:"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   255
      Left            =   6480
      TabIndex        =   8
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblRecord 
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Processing Record:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Directory:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim Result As Long
    'On Error GoTo Err_mnuFileOpen_Click
    CommonDialog1.Filter = "Text File (*.txt) |*.txt| All Files (*.*) |*.*"
    CommonDialog1.Action = 2
    Me.txtDirectory = CommonDialog1.FileName
    Me.Command2.Enabled = True
    Exit Sub
End Sub

Private Sub Command2_Click()
Dim dbs As Database
    Dim rstExport As Recordset
    Dim strCriteria As String
    Dim SQL As String
    Dim intTotalRecord As Integer
    Dim intCount As Integer
    Dim rLine As String
    Dim intPercentHold As Integer
    Dim intPercent As Integer
    Dim intCountHold As Integer
    Dim varRecord As Integer
    
    On Error Resume Next
    Me.Height = 1935
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    
    Set dbs = OpenDatabase(gDatabaseName)
    
    SQL = "Phone Book Table"
    Set rstExport = dbs.OpenRecordset(SQL)
    rstExport.MoveFirst
    intTotalRecord = rstExport.RecordCount
    If Err.Number <> 0 Then
        Module1.errRTN "Dataset Error", "There was a problem opening the Recordset.", Err.Number, Err.Description, 0
        Command1.Enabled = True
        Command2.Enabled = True
        Command3.Enabled = True
        Exit Sub
    End If
    Open Me.txtDirectory For Output As #1 Len = 500
    intCount = 0
    lblPercent = "0%"
    varRecord = 0
    MDIForm1.StatusBar1.Panels(1).Text = "Action: Exporting Data"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: 0%"
    MDIForm1.StatusBar1.Refresh
    Me.Refresh
    Do While intCount < intTotalRecord
        intCount = intCount + 1
        ProgressBar1.Value = (intCount / intTotalRecord) * 100
        lblRecord = intCount & "/" & intTotalRecord
        intPercent = Round(((intCount / intTotalRecord) * 100), 0)
        If intPercent = intPercentHold + 5 Then
            lblPercent = intPercent & "%"
            intPercentHold = intPercent
            MDIForm1.StatusBar1.Panels(2).Text = "Status: " & intPercentHold & "%"
            MDIForm1.StatusBar1.Refresh
            lblPercent.Refresh
        End If
        lblRecord.Refresh

        If Left(rstExport.Fields("FaxNumber"), 3) = Area1.Text Or Left(rstExport.Fields("FaxNumber"), 3) = Area2.Text Or Left(rstExport.Fields("FaxNumber"), 3) = Area3.Text Or Left(rstExport.Fields("FaxNumber"), 3) = Area4.Text Then
            varPhone = "*67," & rstExport.Fields("FaxNumber")
        Else
            varPhone = "*67,1" & rstExport.Fields("FaxNumber")
        End If
        rLine = Chr(34) & rstExport.Fields("FirstName") & " " & rstExport.Fields("LastName") & Chr(34) & "," & Chr(34) & rstExport.Fields("Company") & Chr(34) & "," & Chr(34) & varPhone & Chr(34)
        
        Print #1, rLine
        intCountHold = 0
        If rstExport.Fields("UsedCount") > 0 Then
            intCountHold = rstExport.Fields("UsedCount")
        End If
        With rstExport
            .Edit
            .Fields("UsedCount") = intCountHold + 1
            .Fields("LastUsed") = Date
            .Update
        End With
        rstExport.MoveNext
    Loop
    lblRecord.Refresh
    Close #1
    MsgBox "Export was successful.", vbOKOnly, "Complete"
    Unload Me
End Sub
Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Height = 1245
    Me.Command2.Enabled = False
    MDIForm1.StatusBar1.Panels(1).Text = "Action: Awaiting Export Decision"
    MDIForm1.StatusBar1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.StatusBar1.Panels(1).Text = "Action: None"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: None"
    MDIForm1.StatusBar1.Refresh
End Sub
