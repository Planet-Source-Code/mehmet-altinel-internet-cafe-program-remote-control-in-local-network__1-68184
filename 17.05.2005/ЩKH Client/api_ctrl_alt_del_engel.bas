Attribute VB_Name = "api_ctrl_alt_del_engel"
Public Const HKEY_CURRENT_USER = &H80000001
Public Const REG_SZ = 1
'Reg Aç
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'Reg Kapat
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
'Reg Yaz
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'Reg Sil
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
'98
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public Function TaskMgr(Enabled As Boolean)
Dim Hkey As Long
RegOpenKey HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", Hkey
If Enabled = False Then
RegSetValueEx Hkey, "DisableTaskMgr", 0, REG_SZ, ByVal "0", 1
SystemParametersInfo 97, True, 1, 0
Else
RegDeleteValue Hkey, "DisableTaskMgr"
SystemParametersInfo 97, False, 1, 0
End If
RegCloseKey (Hkey)
End Function

'98 için kod xp de çalýþmaz
'**************************************************************************************************************************************************************************************************************
'Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction _
'As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
'Dim topon As Boolean
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Public Sub DisableCtrlAltDelete(bDisabled As Boolean)
'Dim x As Long
'x = SystemParametersInfo(97, bDisabled, CStr(1), 0)
'End Sub
'*************************************************************************************************************************************************************************************************************
