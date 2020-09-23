VERSION 5.00
Begin VB.Form frmErrRTN 
   Caption         =   "Unknown Error"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   Icon            =   "frmErrRTN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtErrDesc 
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   7335
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   6240
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label txtErrNum 
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Please take down this information and email it to joel@am-pm.com."
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Label Label1 
      Caption         =   "Error Number:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.Label txtMsg 
      Alignment       =   2  'Center
      Caption         =   "An unknown error occured.  Please let us know what you where doing when the error occured."
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmErrRTN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
