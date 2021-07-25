Attribute VB_Name = "modCreateHook"
Option Explicit

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const WH_CBT As Long = 5, HCBT_CREATEWND As Long = 3
Public hCreateHook As Long, hComboLBox As Long
 
Function CreateHookProcForVBCombo(ByVal uCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Const GWL_STYLE = -16, CBS_OWNERDRAWFIXED = &H10, CBS_DROPDOWNLIST = &H3, CBS_HASSTRINGS = &H200
 
  If uCode = HCBT_CREATEWND Then
    Dim Class As String: Class = Space(256)
        Class = Left(Class, GetClassName(wParam, Class, 255))
     If Class = "ThunderComboBox" Or Class = "ThunderRT6ComboBox" Then
       SetWindowLong wParam, GWL_STYLE, (GetWindowLong(wParam, GWL_STYLE) Or CBS_OWNERDRAWFIXED Or CBS_DROPDOWNLIST) And Not CBS_HASSTRINGS
     ElseIf Class = "ComboLBox" Then
       hComboLBox = wParam
     End If
  End If
  
  CreateHookProcForVBCombo = CallNextHookEx(hCreateHook, uCode, wParam, ByVal lParam)
End Function
