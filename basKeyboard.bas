Attribute VB_Name = "Keyboardhandler"
' Keyboardhandler.bas - Demonstrates low-level keyboard hooks
' Copyright (c) 2002. All Rights Reserved
' By Paul Kimmel. pkimmel@softconcepts.com

'http://msdn.microsoft.com/library/default.asp?url=
'/library/en-us/winui/WinUI/WindowsUserInterface
'/Windowing/Hooks/HookReference/HookFunctions/LowLevelKeyboardProc.asp

Option Explicit

Public Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long

Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
  
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cb As Long)

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long

Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long

Private Type KBDLLHOOKSTRUCT
  vkCode As Long
  scanCode As Long
  flags As Long
  time As Long
  dwExtraInfo As Long
End Type

' Low-Level Keyboard Constants
Private Const HC_ACTION = 0
Private Const LLKHF_EXTENDED = &H1
Private Const LLKHF_INJECTED = &H10
Private Const LLKHF_ALTDOWN = &H20
Private Const LLKHF_UP = &H80

' Virtual Keys
Public Const VK_TAB = &H9
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const VK_DELETE = &H2E

Private Const WH_KEYBOARD_LL = 13&
Public KeyboardHandle As Long
Dim CurrentPID As Long

'Public KeyboardHook As KeyboardHook


' Implement this function to block as many key combinations as
' you'd like
Public Function IsHooked(ByRef Hookstruct As KBDLLHOOKSTRUCT) As Boolean
Dim pid As Long
'  If (KeyboardHook Is Nothing) Then
'    IsHooked = True
'    Exit Function
'  End If
  
  If (Hookstruct.vkCode = 78) And CBool(GetAsyncKeyState(VK_CONTROL) And &H8000) Then
    Dim FGwin As Long
    GetWindowThreadProcessId FrmMain.hwnd, pid
    CurrentPID = GetCurrentProcessId
    
    If pid = CurrentPID Then
        IsHooked = True
    Else
        IsHooked = False
    End If

    Exit Function
  End If
  
End Function

Public Function KeyboardCallback(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Static Hookstruct As KBDLLHOOKSTRUCT

  If (Code = HC_ACTION) Then
    ' Copy the keyboard data out of the lParam (which is a pointer)
    Call CopyMemory(Hookstruct, ByVal lParam, Len(Hookstruct))

    If (IsHooked(Hookstruct)) Then
      KeyboardCallback = 1
      Exit Function
    End If

  End If

  KeyboardCallback = CallNextHookEx(KeyboardHandle, Code, wParam, lParam)

End Function

Public Sub HookKeyboard()
  KeyboardHandle = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyboardCallback, App.hInstance, 0&)
'  Call CheckHooked
End Sub

'Public Sub CheckHooked()
'  If (Hooked) Then
'    Debug.Print "Keyboard hooked"
'  Else
'    Debug.Print "Keyboard hook failed: " & Err.LastDllError
'  End If
'End Sub

Private Function Hooked()
  Hooked = KeyboardHandle <> 0
End Function

Public Sub UnhookKeyboard()
  If (Hooked) Then
    Call UnhookWindowsHookEx(KeyboardHandle)
  End If
End Sub
