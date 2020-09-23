VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmailSearch 
   Caption         =   "Change Email Search"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   Icon            =   "frmEmailSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1845
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      Height          =   375
      Left            =   2760
      MaskColor       =   &H00808080&
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Add New"
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.OptionButton SearchBy 
      Caption         =   "Company"
      Height          =   255
      Index           =   2
      Left            =   420
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton SearchBy 
      Caption         =   "First Name"
      Height          =   255
      Index           =   0
      Left            =   420
      TabIndex        =   4
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton SearchBy 
      Caption         =   "Last Name"
      Height          =   255
      Index           =   1
      Left            =   2580
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.OptionButton SearchBy 
      Caption         =   "Fax Number"
      Height          =   255
      Index           =   3
      Left            =   2580
      TabIndex        =   2
      Top             =   720
      Value           =   -1  'True
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Search String:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      Caption         =   "0%"
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
End
Attribute VB_Name = "frmEmailSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    
    If Len(txtSearch.Text) > 0 And Left(txtSearch.Text, 1) <> "_" Then
        frmEmailResults.Show 0, MDIForm1
    Else
        Module1.ErrRTN "Data format error", "You must enter a phone number before continuing.", Err.Number, Err.Description, 0
        Exit Sub
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Height = 1575
End Sub
