VERSION 5.00
Begin VB.Form frmUserReport 
   Caption         =   "Changes by User"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmUserReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox comUser 
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&View Report"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Select User:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmUserReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'Dim varFilePath As String
    'Dim vbWord As Object
    
    'Set vbWord = CreateObject("Word.Application")
    'varFilePath = App.Path & "\Userlo~1.doc"
    
    'Documents(varFilePath).Activate
    
End Sub
