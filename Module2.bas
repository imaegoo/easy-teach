Attribute VB_Name = "GetWebBrowserHwnd"
Option Explicit
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Function GetWebBrowserHwnd(hwndBrowserContainer As Long) As Long
Dim RetVal As Long
Dim hwndPeer As Long
Dim ClassString As String * 256
hwndPeer = GetWindow(hwndBrowserContainer, GW_CHILD)
Do While hwndPeer <> 0
RetVal = GetClassName(hwndPeer, ClassString, 256)
If InStr(ClassString, "Shell Embedding") <> 0 Then
GetWebBrowserHwnd = hwndPeer
Exit Do
End If
hwndPeer = GetWindow(hwndPeer, GW_HWNDNEXT)
Loop
End Function

