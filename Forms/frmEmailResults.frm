VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmailResults 
   Caption         =   "Email Results (Add/Edit)"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7725
   Icon            =   "frmEmailResults.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   6360
      TabIndex        =   1
      Top             =   4680
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstGrid 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7435
      SortKey         =   4
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Company"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "First"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Last"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Fax Number"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Email"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select the record you want to edit from the list bellow.  Do this by double clicking on the record."
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmEmailResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim dbs As Database
    Dim rstSearch As Recordset
    Dim MySQL As String
    
    On Error Resume Next
    txtSearch.Text = frmEmailSearch.txtSearch.Text
    Unload frmEmailSearch
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmEmailSearch.Show 0, MDIForm1
End Sub
