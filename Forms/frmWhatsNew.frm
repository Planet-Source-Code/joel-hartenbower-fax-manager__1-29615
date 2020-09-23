VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWhatsNew 
   Caption         =   "What's New"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   Icon            =   "frmWhatsNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   3105
      TabIndex        =   1
      Top             =   4800
      Width           =   1695
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8070
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmWhatsNew.frx":0742
   End
   Begin VB.Label txtBuild1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Build:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label txtBuild 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Build:"
      Height          =   255
      Left            =   4920
      TabIndex        =   2
      Top             =   4800
      Width           =   2775
   End
End
Attribute VB_Name = "frmWhatsNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    RichTextBox1.FileName = gApp_Path & "WhatsNew.rtf"
    txtBuild1 = "Version: " & App.Major & ".0" & App.Minor & " Build: " & App.Revision
    txtBuild = "Version: " & App.Major & ".0" & App.Minor & " Build: " & App.Revision
End Sub
