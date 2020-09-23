Attribute VB_Name = "Module1"
Global gSQL As String
Global gTitle As String
Global gMsg As String
Global gErrNumber As String
Global gErrDescription As String
Global gstatus As String
Global gDatabaseName As String
Global gUserDatabase As String
Global gExportDB As String
Global gUserName As String
Global gRecordID As String
Global gSearch As String
Global gWhatsNew As String
Global gApp_Path As String
Global gBuildNumber As String

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
(ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal szData As String, ByRef lpcbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Sub DeCryptIt()
    Dim t As String, u As String
    Dim i As Integer, C As Integer
    'gCompanyName = EnCryption(gCompanyName)
    'gCompanyAddr1 = EnCryption(gCompanyAddr1)
    'gCompanyAddr2 = EnCryption(gCompanyAddr2)
    'gCompanyCity = EnCryption(gCompanyCity)
    'gCompanyState = EnCryption(gCompanyState)
    'gCompanyZip = EnCryption(gCompanyZip)
    'gAccountNumber = EnCryption(gAccountNumber)
End Sub
Function EnCryption(MyString As String) As String
    Dim t As String, u As String
    Dim i As Integer, C As Integer
    t = MyString
    u = ""
    For i = 1 To Len(t)
        C = Asc(Mid$(t, i, 1))
        u = u + Chr$(C Xor &HFF)
    Next
    EnCryption = u

End Function

Public Function IniRead(FileName As String, Section As String, Key As String)
    Dim Result, temp As String
    temp = Space(255)
    Result = GetPrivateProfileString(Section, Key, "", temp, Len(temp), FileName)
    If Val(Result) = 0 Then
        IniRead = ""
    Else
        temp = Trim$(temp)
        If Asc(Right$(temp, 1)) = 0 Then
            temp = Left$(temp, Len(temp) - 1)
        End If
        IniRead = temp

    End If
End Function

Public Function IniWrite(fname As String, Sect As String, Key As String, Value As String)
    IniWrite = WritePrivateProfileString(Sect, Key, Value, fname)
End Function
'
'-----------------------------------------------------------
' SUB: AddDirSep
' Add a trailing directory path separator (back slash) to the
' end of a pathname unless one already exists
'
' IN/OUT: [strPathName] - path to add separator to
'-----------------------------------------------------------
'
Sub AddDirSep(strPathName As String)
    If Right$(RTrim$(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub
'-----------------------------------------------------------
' FUNCTION: GetWindowsDir
'
' Calls the windows API to get the windows directory and
' ensures that a trailing dir separator is present
'
' Returns: The windows directory
'-----------------------------------------------------------
'
Function GetWindowsDir() As String
#If Win16 Then
wpName = gstrWP16
#End If

#If Win32 Then
wpName = gstrWP32
#End If

    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    'If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
    '    strBuf = StripTerminator$(strBuf)
    '    AddDirSep strBuf

    '    GetWindowsDir = UCase16(strBuf)
    'Else
    '    GetWindowsDir = gstrNULL
    'End If
End Function
Public Sub CenterForm(myForm As Form)
    myForm.Left = (Screen.Width - myForm.Width) / 2
    myForm.Top = (Screen.Height - myForm.Height) / 2
End Sub

Function app_path() As String
    x = App.Path
    If Right$(x, 1) <> "\" Then x = x + "\"
    gApp_Path = UCase$(x)
End Function
Function errRTN()
    Load frmErrRTN
    frmErrRTN.Caption = gTitle
    frmErrRTN.txtMsg.Caption = gMsg
    frmErrRTN.txtErrNum.Caption = gErrNumber
    frmErrRTN.txtErrDesc.Text = gErrDescription
    
    If gstatus = 0 Then
        frmErrRTN.cmdOK.Visible = True
        frmErrRTN.cmdQuit.Visible = False
        frmErrRTN.Label2.Visible = False
    Else
        frmErrRTN.cmdOK.Visible = False
        frmErrRTN.cmdQuit.Visible = True
        frmErrRTN.Label2.Visible = True
    End If
    frmErrRTN.Show vbModal
End Function
Function GetRegValue(hKey As Long, lpszSubKey As String, _
    szKey As String, szDefault As String) As Variant

    On Error GoTo ErrorRoutineErr:
    
    Dim phkResult As Long
    Dim lResult As Long
    Dim szBuffer As String
    Dim lBuffSize As Long
    
    'Create Buffer
    szBuffer = Space(255)
    lBuffSize = Len(szBuffer)
    
    'Open the key
    RegOpenKeyEx hKey, lpszSubKey, 0, 1, phkResult
    
    'Query the value
    lResult = RegQueryValueEx(phkResult, szKey, 0, _
        0, szBuffer, lBuffSize)
    
    'Close the key
    RegCloseKey phkResult
    
    'Return obtained value
    If lResult = ERROR_SUCCESS Then
        GetRegValue = Left(szBuffer, lBuffSize - 1)
    Else
        GetRegValue = szDefault
    End If
    Exit Function
    
ErrorRoutineErr::
    MsgBox "ERROR #" & Str$(Err) & " : " & Error & Chr(13) _
         & "Please exit and try again."
    GetRegValue = ""

End Function

