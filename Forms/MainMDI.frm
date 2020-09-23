VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Fax Manager v3.0 (Alpha)"
   ClientHeight    =   5565
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7605
   Icon            =   "MainMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMDI.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMDI.frx":0B94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMDI.frx":0FE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMDI.frx":1438
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMDI.frx":188A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMDI.frx":1CDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMDI.frx":212E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainMDI.frx":2580
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Description     =   "Open Database"
            Object.ToolTipText     =   "Open Database"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Close"
            Description     =   "Close Database"
            Object.ToolTipText     =   "Close Database"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Add"
            Description     =   "Add/Edit/Delete"
            Object.ToolTipText     =   "Add/Edit/Delete"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Description     =   "Quick Delete"
            Object.ToolTipText     =   "Quick Delete"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Export"
            Description     =   "Export Phone Book"
            Object.ToolTipText     =   "Export Phone Book"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Description     =   "What's New?"
            Object.ToolTipText     =   "What's New?"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Description     =   "Exit Fax Manager"
            Object.ToolTipText     =   "Exit Fax Manager"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   5310
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1799
            MinWidth        =   1058
            Text            =   "Status: None"
            TextSave        =   "Status: None"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1799
            MinWidth        =   882
            Text            =   "Action: None"
            TextSave        =   "Action: None"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   0
            Text            =   "Open DB: "
            TextSave        =   "Open DB: "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1085
            MinWidth        =   882
            Text            =   "Version"
            TextSave        =   "Version"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2275
            MinWidth        =   1058
            Text            =   "Username: None"
            TextSave        =   "Username: None"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1376
            MinWidth        =   882
            TextSave        =   "10:55 AM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1535
            MinWidth        =   882
            TextSave        =   "12/9/2001"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc datFaxDB 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   7605
      _ExtentX        =   13414
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
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Open"
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Close"
         Index           =   1
      End
      Begin VB.Menu mnuFileItem 
         Caption         =   "&Exit"
         Index           =   2
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuMaintenece 
      Caption         =   "&Maintenece"
      Begin VB.Menu mnuMainteneceItem 
         Caption         =   "&Add/Edit/Delete"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuMainteneceItem 
         Caption         =   "&Email Change"
         Index           =   1
         Shortcut        =   ^E
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainteneceItem 
         Caption         =   "&Quick Delete"
         Index           =   2
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportItem 
         Caption         =   "&Changes by User"
         Index           =   0
      End
   End
   Begin VB.Menu mnuMisc 
      Caption         =   "&Miscellaneous"
      Begin VB.Menu mnuMiscItem 
         Caption         =   "Export - &Access"
         Index           =   0
      End
      Begin VB.Menu mnuMiscItem 
         Caption         =   "Export - &Text"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMiscItem 
         Caption         =   "&Check for Duplicates"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMiscItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuMiscItem 
         Caption         =   "Change Password"
         Index           =   4
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Administrator"
      Begin VB.Menu mnuAdminItem 
         Caption         =   "&User Editor"
         Index           =   0
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&Info"
      Begin VB.Menu mnuInfoItem 
         Caption         =   "&What's New"
         Index           =   0
      End
      Begin VB.Menu mnuInfoItem 
         Caption         =   "&Instructions"
         Index           =   1
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public varErrRTN As Integer

Private Sub MDIForm_Load()
    MDIForm1.StatusBar1.Panels(4).Text = "v" & App.Major & ".0" & App.Minor & " Build: " & App.Revision & " "
    If gDatabaseName = "Not Found" Then
        MDIForm1.StatusBar1.Panels(3).Text = "Open DB: -DATABASE CLOSED-"
        MDIForm1.mnuMaintenece.Enabled = False
        MDIForm1.mnuMiscItem(0).Enabled = False
        MDIForm1.mnuFileItem(1).Enabled = False
        MDIForm1.mnuFileItem(0).Enabled = True
        MDIForm1.mnuAdmin.Visible = False
        MDIForm1.Toolbar1.Buttons(1).Enabled = True
        MDIForm1.Toolbar1.Buttons(2).Enabled = False
        MDIForm1.Toolbar1.Buttons(4).Enabled = False
        MDIForm1.Toolbar1.Buttons(5).Enabled = False
        MDIForm1.Toolbar1.Buttons(6).Enabled = False
    Else
        CheckFax
    End If
End Sub

Private Sub mnuAdminItem_Click(Index As Integer)
    Select Case Index
        Case 0
            frmUserAdmin.Show 0, MDIForm1
            Exit Sub
    End Select
End Sub

Private Sub mnuFileItem_Click(Index As Integer)
    Select Case Index
        Case 0
            OpenDB
            Exit Sub
        Case 1
            CloseDB
            Exit Sub
        Case 2
            End
    End Select
End Sub

Sub OpenDB()
    Dim strCriteria, varUserName, varDBName, varDBNameHold As String
    Dim Result As Long
    
    CommonDialog1.CancelError = True
    On Error GoTo errRoutine
    
    CommonDialog1.Filter = "Database (*.mdb) |*.mdb| All Files (*.*) |*.*"
    Me.MousePointer = 11
    CommonDialog1.Action = 1
    Me.MousePointer = 0
    gDatabaseName = CommonDialog1.FileName
    If Err.Number <> 3059 Then
        Dim dbsRegister As Database
        Dim strAttributes As String
        
        'Register the new database for ODBC
        gTitle = "Registery Error"
        gMsg = "Could not register the database in the system."
        gstatus = 1
        strAttributes = "DBQ=" & gDatabaseName
        DBEngine.RegisterDatabase "FaxManagerDB", "Microsoft Access Driver (*.mdb)", True, strAttributes
        CheckFax
        
        If varErrRTN = 1 Then
            CloseDB
            Exit Sub
        End If
        
        varCount = 0
        Do While Left(varDBNameHold, 1) <> "\" And varCount < Len(gDatabaseName)
            varDBNameHold = Right(gDatabaseName, 34 - varCount)
            varCount = varCount + 1
        Loop
        varDBName = Left(varDBNameHold, Len(varDBNameHold) - 4)
        varDBName = Right(varDBName, Len(varDBName) - 1)
        MDIForm1.StatusBar1.Panels(3).Text = "Open DB: " & varDBName
        MDIForm1.StatusBar1.Refresh
        
        gTitle = "Recordset Error"
        gMsg = "There was a problem opening the recordset for a save."
        gstatus = 0
        varUserName = EnCryption(gUserName)
        EnableItems
    End If
    Exit Sub

errRoutine:
    If Err.Number = 3059 Or Err.Number = 32755 Then
        Me.MousePointer = 0
        Exit Sub
    End If
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
End Sub

Private Sub mnuInfoItem_Click(Index As Integer)
    Select Case Index
        Case 0
            frmWhatsNew.Show 1, MDIForm1
            Exit Sub
        Case 1
            frmInstructions.Show 0, MDIForm1
            Exit Sub
    End Select
End Sub

Private Sub mnuMainteneceItem_Click(Index As Integer)
    Select Case Index
        Case 0
            frmSearch.Show 0, MDIForm1
            Exit Sub
        Case 1
            frmEmailSearch.Show 0, MDIForm1
            Exit Sub
        Case 2
            frmQuickDelSearch.Show 0, MDIForm1
            Exit Sub
    End Select
End Sub

Private Sub mnuMiscItem_Click(Index As Integer)
    Select Case Index
        Case 0
            frmWinFax.Show 0, MDIForm1
            Exit Sub
        Case 1
            frmExport.Show 0, MDIForm1
            Exit Sub
        Case 2
            frmDuplicate.Show
            Exit Sub
        Case 4
            frmChangePW.Show vbModal, MDIForm1
            Exit Sub
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Open"
            OpenDB
            Exit Sub
        Case "Close"
            CloseDB
            Exit Sub
        Case "Add"
            frmSearch.Show 0, MDIForm1
            Exit Sub
        Case "Delete"
            frmQuickDelSearch.Show 0, MDIForm1
            Exit Sub
        Case "Export"
            frmWinFax.Show 0, MDIForm1
            Exit Sub
        Case "New"
            frmWhatsNew.Show 1, MDIForm1
            Exit Sub
        Case "Exit"
            End
    End Select
End Sub
Sub EnableItems()
    MDIForm1.mnuFileItem(1).Enabled = True
    MDIForm1.mnuFileItem(0).Enabled = False
    MDIForm1.mnuMaintenece.Enabled = True
    MDIForm1.mnuReports.Enabled = True
    MDIForm1.mnuMisc.Enabled = True
    MDIForm1.mnuMiscItem(0).Enabled = True
    MDIForm1.Toolbar1.Buttons(1).Enabled = False
    MDIForm1.Toolbar1.Buttons(2).Enabled = True
    MDIForm1.Toolbar1.Buttons(4).Enabled = True
    MDIForm1.Toolbar1.Buttons(5).Enabled = True
    MDIForm1.Toolbar1.Buttons(6).Enabled = True
    MDIForm1.Toolbar1.Buttons(7).Enabled = True
End Sub
Sub DisableItems()
    MDIForm1.mnuFileItem(1).Enabled = False
    MDIForm1.mnuFileItem(0).Enabled = True
    MDIForm1.mnuMaintenece.Enabled = False
    MDIForm1.mnuReports.Enabled = False
    MDIForm1.mnuMisc.Enabled = False
    MDIForm1.Toolbar1.Buttons(1).Enabled = True
    MDIForm1.Toolbar1.Buttons(2).Enabled = False
    MDIForm1.Toolbar1.Buttons(4).Enabled = False
    MDIForm1.Toolbar1.Buttons(5).Enabled = False
    MDIForm1.Toolbar1.Buttons(6).Enabled = False
    MDIForm1.Toolbar1.Buttons(7).Enabled = False
End Sub
Sub errRoutine()
    If gstatus <> "1" Or gstatus <> "0" Then
        gstatus = 0
    End If
    Module1.errRTN
End Sub
Sub CloseDB()
    Dim strCriteria, varUserName As String
    Dim Result As Long
    Dim hKey As Long
    Dim strPath As String
    
    On Error Resume Next
    
    'CloseForms
    'Deleting Previous Open DB Information
    On Error GoTo errRoutine
    gTitle = "Registery Error"
    gMsg = "There was an error deleting the registry key."
    gstatus = 0
    gDatabaseName = ""
    strAttributes = "DBQ=" & gDatabaseName
    DBEngine.RegisterDatabase "FaxManagerDB", "Microsoft Access Driver (*.mdb)", True, strAttributes
    MDIForm1.StatusBar1.Panels(3).Text = "Open DB: -- No Database Selected --"
    MDIForm1.StatusBar1.Refresh
    CloseForms
    DisableItems
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
End Sub
Sub CheckFax()
    Dim rstPhone As Recordset
    
    On Error GoTo errRoutine
    
    gTitle = "Database Error"
    gMsg = "Could not register the database in the system."
    gstatus = 1
    
    gSQL = "SELECT * FROM [Phone Book Table]"
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh
    
    strCriteria = Left(datFaxDB.Recordset("FaxNumber"), 4)
    If strCriteria = "*67," Then
        strCriteria = MsgBox("The fax numbers in this database need to be fixed, would you like to do so now?", vbYesNo, "Invalid format")
        If strCriteria = vbYes Then
            frmTrimFax.Show 0, MDIForm1
            Exit Sub
        End If
    End If
    Exit Sub
    
errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    varErrRTN = 1
    errRoutine
End Sub
Sub CloseForms()
    On Error Resume Next
    
    Unload frmAddEdit
    Unload frmChangePW
    Unload frmEmailResults
    Unload frmEmailSearch
    Unload frmErrRTN
    Unload frmExport
    Unload frmInstructions
    Unload frmQuickDelSearch
    Unload frmQuickDelVerify
    Unload frmResultsEdit
    Unload frmSearch
    Unload frmUserReport
    Unload frmWhatsNew
    Unload frmWinFax
End Sub
