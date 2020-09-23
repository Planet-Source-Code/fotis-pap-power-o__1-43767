Attribute VB_Name = "general"
Option Explicit
Public easYmove As Boolean
Public hiDDen As Boolean
Public Declare Function SetWindowLong Lib "user32" Alias _
   "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
 
Public Declare Function CallWindowProc Lib "user32" _
   Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
   ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, _
   ByVal lParam As Long) As Long

Public Const GWL_WNDPROC = (-4)

Public Const WM_CONTEXTMENU = &H7B

Global lpPrevWndProc As Long
Global gHW As Long

Public Sub Hook()
   lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, _
   AddressOf gWindowProc)
End Sub

Public Sub Unhook()
   Dim temp As Long
   temp = SetWindowLong(gHW, GWL_WNDPROC, lpPrevWndProc)
End Sub

Public Function gWindowProc(ByVal hWnd As Long, ByVal Msg As Long, _
                 ByVal wParam As Long, ByVal lParam As Long) As Long
   If Msg = WM_CONTEXTMENU Then
      Debug.Print "Intercepted WM_CONTEXTMENU at " & Now
      gWindowProc = True
   Else ' Send all other messages to the default message handler
      gWindowProc = CallWindowProc(lpPrevWndProc, hWnd, Msg, wParam, _
         lParam)
   End If
End Function


