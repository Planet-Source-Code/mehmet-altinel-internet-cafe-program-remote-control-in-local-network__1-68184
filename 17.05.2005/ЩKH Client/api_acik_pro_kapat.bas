Attribute VB_Name = "api_acik_pro_kapat"
Option Explicit

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_CLOSE = &H10


Public Function KillApp(ByVal WinAppName As String) As Boolean
    Dim winHwnd As Long
    Dim RetVal As Long
    
    winHwnd = FindWindow(vbNullString, WinAppName)
    
    Debug.Print winHwnd
    
    If winHwnd <> 0 Then
        RetVal = PostMessage(winHwnd, WM_CLOSE, 0&, 0&)
        If RetVal = 0 Then
            KillApp = False
            Exit Function
        End If
    Else
        KillApp = False
        Exit Function
    End If
    
    KillApp = True
End Function
