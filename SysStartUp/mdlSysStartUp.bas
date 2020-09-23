Attribute VB_Name = "mdlSysStartUp"

'------------------------------------------------------------------------------------------------
' Auther    : Jim Jose
' email     : jimjosev33@yahoo.com
' Credits   : NatureOfOmega, for the basics and Local machine path
' Purpose   : Make your software auto-executable on system startup.
'           : Code can be used to add any file to system's Startup item list
'------------------------------------------------------------------------------------------------

Option Explicit

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_USERS = &H80000003
Private Const ERROR_SUCCESS = 0&
Private Const REG_DWORD = 4
Private Const REG_SZ = 1

'----------------------------------------------------------------------------------------
' Syntax    : AddToStartUp App.EXEName, App.Path & "\" & App.EXEName, True
'----------------------------------------------------------------------------------------
'  mItemKey : Just a key to register
'  mPath    : The path of the file that to be executed
'  mState   : Sets the value.. True = [Enabled, loads on startup}
'----------------------------------------------------------------------------------------

Public Function AddToStartUp(mItemKey As String, mPath As String, ByVal mState As Boolean) As Boolean
Dim keyhand As Long
Dim Rtn As Long
Const StrPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"

    On Error GoTo Handle
    Rtn = RegCreateKey(HKEY_LOCAL_MACHINE, StrPath, keyhand)
    If mState Then
        Rtn = RegSetValueEx(keyhand, mItemKey, 0, REG_SZ, ByVal mPath, Len(mPath))
    Else
        Rtn = RegDeleteValue(keyhand, mItemKey)
    End If
    Rtn = RegCloseKey(keyhand)
    
    AddToStartUp = True
    
Exit Function
Handle:
    AddToStartUp = False
End Function

