VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmResultsEdit 
   AutoRedraw      =   -1  'True
   Caption         =   "Search Results for ADD/EDIT"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   Icon            =   "frmResultsEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc datFaxDB 
      Height          =   330
      Left            =   240
      Top             =   4680
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
   Begin MSComctlLib.ListView lstGrid 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   480
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
         Text            =   "Deleted"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   255
      Left            =   6360
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Select the record you want to edit from the list bellow.  Do this by double clicking on the record."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmResultsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim itmX As ListItem
    Dim colX As ColumnHeader
    Dim strField As String
    Dim strItem1, strItem2, strItem3, strItem4, strHold As String
    Dim intTotalRecords, intPercent, intPercentHold As Long
    Dim intCount As Long
    Dim varOrder As String
    Dim varSearch As String
    
    MDIForm1.StatusBar1.Panels(1).Text = "Action: Awaiting Record Selection"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: None"
    MDIForm1.StatusBar1.Refresh
    
    On Error GoTo errRoutine
    strSearch = frmSearch.txtSearch
    If frmSearch.SearchBy(0).Value = True Then varSearch = "FirstName Like '" & strSearch & "'"
    If frmSearch.SearchBy(1).Value = True Then varSearch = "LastName Like '" & strSearch & "'"
    If frmSearch.SearchBy(2).Value = True Then varSearch = "Company Like '" & strSearch & "'"
    If frmSearch.SearchBy(3).Value = True Then varSearch = "FaxNumber Like '" & strSearch & "'"

    gSQL = "SELECT * FROM [Deleted Phone Book Table]"
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh
    
    datFaxDB.Recordset.Filter = varSearch
    intTotalRecord = datFaxDB.Recordset.RecordCount
    
    gSQL = "SELECT * FROM [Phone Book Table]"
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh
    
    datFaxDB.Recordset.Filter = varSearch
    intTotalRecord = intTotalRecord + datFaxDB.Recordset.RecordCount
    
    frmSearch.lblPercent = "0%"
    MDIForm1.StatusBar1.Panels(1).Text = "Action: Searching"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: 0%"
    MDIForm1.StatusBar1.Refresh
    frmSearch.lblPercent.Refresh
    
    lstGrid.Sorted = True
    lstGrid.SortOrder = lvwAscending
    If frmSearch.SearchBy(0).Value = True Then
        strField = "FirstName"
        varOrder = " ORDER BY FirstName"
    End If
    If frmSearch.SearchBy(1).Value = True Then
        strField = "LastName"
        varOrder = " ORDER BY LastName"
    End If
    If frmSearch.SearchBy(2).Value = True Then
        strField = "Company"
        varOrder = " ORDER BY Company"
    End If
    If frmSearch.SearchBy(3).Value = True Then
        strField = "FaxNumber"
        varOrder = " ORDER BY FaxNumber"
    End If
    
    gSQL = "SELECT * FROM [Phone Book Table]" & varOrder
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh
    datFaxDB.Recordset.Filter = varSearch
    
    intCount = 0
    frmSearch.ProgressBar1.Visible = True
    frmSearch.lblPercent.Visible = True
    frmSearch.lblWait.Visible = False
    frmSearch.Refresh
    If Not datFaxDB.Recordset.EOF Then
        Do While datFaxDB.Recordset.EOF = False
            intCount = intCount + 1
            frmSearch.ProgressBar1.Value = (intCount / intTotalRecord) * 100
            intPercent = Round(((intCount / intTotalRecord) * 100), 0)
            If intPercent = intPercentHold + 5 Then
                frmSearch.lblPercent = intPercent & "%"
                intPercentHold = intPercent
                MDIForm1.StatusBar1.Panels(2).Text = "Status: " & intPercentHold & "%"
                MDIForm1.StatusBar1.Refresh
                frmSearch.lblPercent.Refresh
            End If
            If intPercent = "90" Or intPercentHold = 100 Then
                frmSearch.ProgressBar1.Value = 100
                frmSearch.lblPercent = "100%"
                MDIForm1.StatusBar1.Panels(2).Text = "Status: 100%"
                MDIForm1.StatusBar1.Refresh
                frmSearch.lblPercent.Refresh
                intPercentHold = 100
            End If
            
            If Len(datFaxDB.Recordset("Company")) > 0 Then
                strItem1 = datFaxDB.Recordset("Company")
            Else
                strItem1 = " "
            End If
            If Len(datFaxDB.Recordset("FirstName")) > 0 Then
                strItem2 = datFaxDB.Recordset("FirstName")
            Else
                strItem2 = " "
            End If
            If Len(datFaxDB.Recordset("LastName")) > 0 Then
                strItem3 = datFaxDB.Recordset("LastName")
            Else
                strItem3 = " "
            End If
            If Len(datFaxDB.Recordset("FaxNumber")) > 0 Then
                strHold = Right(datFaxDB.Recordset("FaxNumber"), 7)
                strItem4 = "(" & Left(datFaxDB.Recordset("FaxNumber"), 3) & ") " & Left(strHold, 3) & "-" & Right(datFaxDB.Recordset("FaxNumber"), 4)
            Else
                strItem4 = " "
            End If
            Set itmX = lstGrid.ListItems.Add(, , strItem1)
            itmX.SubItems(1) = strItem2
            itmX.SubItems(2) = strItem3
            itmX.SubItems(3) = strItem4
            itmX.SubItems(4) = "No"
            datFaxDB.Recordset.MoveNext
        Loop
        gSQL = "SELECT * FROM [Deleted Phone Book Table]"
        datFaxDB.RecordSource = gSQL
        datFaxDB.Refresh
    
        datFaxDB.Recordset.Filter = varSearch
        If Not datFaxDB.Recordset.EOF Then
            Do While datFaxDB.Recordset.EOF = False
                intCount = intCount + 1
                frmSearch.ProgressBar1.Value = (intCount / intTotalRecord) * 100
                intPercent = Round(((intCount / intTotalRecord) * 100), 0)
                frmSearch.lblPercent = intPercent & "%"
                intPercentHold = intPercent
                MDIForm1.StatusBar1.Panels(2).Text = "Status: " & intPercentHold & "%"
                MDIForm1.StatusBar1.Refresh
                frmSearch.lblPercent.Refresh
                If Len(datFaxDB.Recordset("Company")) > 0 Then
                    strItem1 = datFaxDB.Recordset("Company")
                Else
                    strItem1 = " "
                End If
                If Len(datFaxDB.Recordset("FirstName")) > 0 Then
                    strItem2 = datFaxDB.Recordset("FirstName")
                Else
                    strItem2 = " "
                End If
                If Len(datFaxDB.Recordset("LastName")) > 0 Then
                    strItem3 = datFaxDB.Recordset("LastName")
                Else
                    strItem3 = " "
                End If
                If Len(datFaxDB.Recordset("FaxNumber")) > 0 Then
                    strHold = Right(datFaxDB.Recordset("FaxNumber"), 7)
                    strItem4 = "(" & Left(datFaxDB.Recordset("FaxNumber"), 3) & ") " & Left(strHold, 3) & "-" & Right(datFaxDB.Recordset("FaxNumber"), 4)
                Else
                    strItem4 = " "
                End If
                'lstGrid.ForeColor = &HFF&
                Set itmX = lstGrid.ListItems.Add(, , strItem1)
                itmX.SubItems(1) = strItem2
                itmX.SubItems(2) = strItem3
                itmX.SubItems(3) = strItem4
                itmX.SubItems(4) = "Yes"
                datFaxDB.Recordset.MoveNext
            Loop
        End If
        MDIForm1.StatusBar1.Panels(1).Text = "Action: Awaiting Record Selection"
        MDIForm1.StatusBar1.Panels(2).Text = "Status: None"
        MDIForm1.StatusBar1.Refresh
    Else
        Unload Me
        Exit Sub
    End If
    Exit Sub

errRoutine:
    gErrNumber = Err.Number
    gErrDescription = Err.Description
    errRoutine
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmSearch.Height = 1950
    frmSearch.Refresh
    MDIForm1.StatusBar1.Panels(1).Text = "Action: None"
    MDIForm1.StatusBar1.Panels(2).Text = "Status: None"
    MDIForm1.StatusBar1.Refresh
End Sub

Private Sub lstGrid_DblClick()
    Dim strSearchText As String
    Dim strHold, strHold1 As String
    
    On Error GoTo errRoutine
    strSearchText = lstGrid.SelectedItem.ListSubItems(3)
    strHold = Left(strSearchText, 4)
    strHold = Right(strHold, 3)
    strHold1 = Right(strSearchText, 8)
    strHold = strHold & Left(strHold1, 3)
    strSearchText = strHold & Right(strHold1, 4)
    
    gSQL = "SELECT * FROM [Phone Book Table] WHERE FaxNumber = '" & strSearchText & "'"
    datFaxDB.RecordSource = gSQL
    datFaxDB.Refresh
    
    If datFaxDB.Recordset.EOF = False Then
        gRecordID = datFaxDB.Recordset("PID")
    Else
        MsgBox "That record was not found or was previously deleted.", vbInformation, "Record Not Found"
        Exit Sub
    End If
    frmAddEdit.Show 1
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

