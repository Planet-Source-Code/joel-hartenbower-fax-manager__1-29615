VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDuplicate 
   Caption         =   "Check for Duplicates"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "&Start"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Done"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<< &Remove"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add >>"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.ListBox lstRemove 
      Height          =   3570
      Left            =   4800
      TabIndex        =   2
      Top             =   840
      Width           =   3495
   End
   Begin VB.ListBox lstRaw 
      Height          =   3570
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3495
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
   Begin MSAdodcLib.Adodc datDupe 
      Height          =   330
      Left            =   2400
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
      Caption         =   "datDupe"
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
   Begin VB.Label Label1 
      Caption         =   $"frmDuplicate.frx":0000
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "frmDuplicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    On Error GoTo errRoutine
    gSQL = "SELECT * FROM [Phone Book Table]"
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh
    
    gSQL = "SELECT * FROM [Duplicate Entries Table]"
    datDupe.RecordSource = gSQL
    datDupe.Refresh
    
    
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
Private Sub Form_Load()
    Me.Height = 5220
    Me.Width = 8520
    
    lstRaw.Enabled = False
    lstRemove.Enabled = False
End Sub
