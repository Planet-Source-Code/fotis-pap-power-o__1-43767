Attribute VB_Name = "internet"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Enum T_WindowStyle
Maximized = 3
Normal = 1
ShowOnly = 5
End Enum
Public Sub OpenInternet(Parent As Form, URL As String, _
WindowStyle As T_WindowStyle)
ShellExecute Parent.hwnd, "Open", URL, "", "", WindowStyle
End Sub


