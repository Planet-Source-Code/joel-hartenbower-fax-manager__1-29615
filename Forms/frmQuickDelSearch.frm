VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmQuickDelSearch 
   Caption         =   "Quick Delete"
   ClientHeight    =   1305
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   Icon            =   "frmQuickDelSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSMask.MaskEdBox txtSearch 
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      AllowPrompt     =   -1  'True
      AutoTab         =   -1  'True
      MaxLength       =   10
      Mask            =   "##########"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Search"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Search:"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter a fax number in the field bellow without any formating (i.e. 9136632451)."
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmQuickDelSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    On Error Resume Next
    frmQuickDelVerify.Show 0, frmQuickDelSearch
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    MDIForm1.StatusBar1.Panels(1).Text = "Action: Awaiting Search Entry"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: None"
    MDIForm1.StatusBar1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIForm1.StatusBar1.Panels(1).Text = "Action: None"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: None"
    MDIForm1.StatusBar1.Refresh
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then frmQuickDelVerify.Show 0, frmQuickDelSearch
End Sub
