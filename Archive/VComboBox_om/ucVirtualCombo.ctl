VERSION 5.00
Begin VB.UserControl ucVirtualCombo 
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   735
   KeyPreview      =   -1  'True
   ScaleHeight     =   630
   ScaleWidth      =   735
End
Attribute VB_Name = "ucVirtualCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Click()
Event RollUp()
Event DropDown()
Event ListMultiClick()
Event KeyPress(KeyAscii As Integer, sKey As String)
Event MouseMoveOnItem(ByVal x As Long, ByVal y As Long)
Event OwnerDraw(ByVal Index As Long, ByVal IsSelected As Boolean, ByVal IsComboItem As Boolean, Canvas As PictureBox, ByVal dx As Long, ByVal dy As Long)

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Type DRAWITEMSTRUCT
  CtlType As Long
  CtlID As Long
  itemID As Long
  itemAction As Long
  itemState As Long
  hwndItem As Long
  hDC As Long
  rcItem As RECT
  ItemData As Long
End Type
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TextOutW& Lib "gdi32" (ByVal hDC&, ByVal x&, ByVal y&, ByVal lpString&, ByVal nCount&)
Private Declare Function GetWindowOrgEx& Lib "gdi32" (ByVal hDC&, lpPoint As Any)
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal enmBar As Long, ByVal bShow As Long) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal enmBar As Long, lpSI As SCROLLINFO) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal enmBar As Long, lpSI As SCROLLINFO, ByVal bRedraw As Long) As Long

Private WithEvents oCB As VB.ComboBox, oBB As VB.PictureBox
Attribute oCB.VB_VarHelpID = -1
Private WithEvents oSL As cSubClass, WithEvents oSC As cSubClass
Attribute oSL.VB_VarHelpID = -1
Attribute oSC.VB_VarHelpID = -1
Private mItemHeight As Long, mMinVisible As Long, hWndLB As Long, mHoverIdx As Long
Public MultiSelect As Boolean, MouseWheelItemScrollDistance As Long

'***** UserControl-EventHandlers
Private Sub UserControl_Initialize()
  Set oBB = Controls.Add("VB.PictureBox", "oBB")
      oBB.AutoRedraw = True: oBB.BorderStyle = 0: oBB.ScaleMode = vbPixels: oBB.BackColor = &H80000005
  hCreateHook = SetWindowsHookEx(WH_CBT, AddressOf modCreateHook.CreateHookProcForVBCombo, 0, App.ThreadID)
    Set oCB = Controls.Add("VB.ComboBox", "oCB")
  UnhookWindowsHookEx hCreateHook
  hWndLB = hComboLBox

  oCB.Visible = True
  ItemHeight = 19
  MinVisibleItems = 10
  MouseWheelItemScrollDistance = 3
End Sub
Private Sub UserControl_Show()
  If Not oSC Is Nothing Then oSC.UnHook
  If Not oSL Is Nothing Then oSL.UnHook
  If Not Ambient.UserMode Then Exit Sub
  Set oSC = New cSubClass: oSC.Hook UserControl.hWnd
  Set oSL = New cSubClass: oSL.Hook hWndLB
End Sub
Private Sub UserControl_Hide()
  If Not oSC Is Nothing Then oSC.UnHook
  If Not oSL Is Nothing Then oSL.UnHook
End Sub
Private Sub UserControl_Resize()
  On Error Resume Next
  oCB.Move 0, 0, UserControl.Width
  UserControl.Height = oCB.Height
End Sub
Private Sub UserControl_EnterFocus()
  oCB.SetFocus
End Sub
Private Sub UserControl_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii, ChrW$(KeyAscii))
  If DroppedDownState And KeyAscii = vbKeySpace And ListIndex > -1 Then
    RaiseEvent ListMultiClick
    Me.Refresh
  End If
End Sub

'***** VB.Combo-EventHandlers
Private Sub oCB_Click()
  RaiseEvent Click
  Me.Refresh
End Sub
Private Sub oCB_DropDown()
  RaiseEvent DropDown
End Sub

'***** SubClassing-EventHandlers
Private Sub oSL_WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, Ret As Long, DefCall As Boolean)
  Const WM_LBUTTONDOWN = &H201, WM_LBUTTONUP = &H202, WM_LBUTTONDBLCLK = &H203
  Const WM_MOUSEMOVE = &H200, WM_MOUSEWHEEL = &H20A, WM_CAPTURECHANGED = &H215
  Const LB_GETCURSEL = &H188, LB_SETCURSEL = &H186, LB_ITEMFROMPOINT = &H1A9 ', LB_GETITEMRECT = &H198
  Static x, y, Rct As RECT, ItemIdx As Long, Closing As Boolean
  
  On Error Resume Next
    Select Case Msg
      Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_LBUTTONDBLCLK
        x = lParam And &HFFFF&: y = lParam \ &H10000
        GetClientRect hWndLB, Rct
        If MultiSelect And x > 0 And x < Rct.Right And y > 0 And y < Rct.Bottom Then
          If Msg = WM_LBUTTONUP And ListIndex > -1 Then RaiseEvent ListMultiClick
          Me.Refresh: DefCall = False
        End If
 
      Case WM_CAPTURECHANGED
        Closing = True
        
      Case WM_MOUSEMOVE
        mHoverIdx = SendMessage(hWnd, LB_ITEMFROMPOINT, 0, ByVal lParam)
        If mHoverIdx > -1 And mHoverIdx < ListCount Then RaiseEvent MouseMoveOnItem(lParam Mod 65536, (lParam \ 65536) Mod mItemHeight)
      Case WM_MOUSEWHEEL
        Rct.Left = lParam Mod 65536: Rct.Top = lParam \ 65536
        ScreenToClient hWnd, Rct
        If oCB.TopIndex - Sgn(wParam And &HFFFF) * MouseWheelItemScrollDistance < 0 Then
          oCB.TopIndex = 0
        Else
          oCB.TopIndex = oCB.TopIndex - Sgn(wParam And &HFFFF) * MouseWheelItemScrollDistance
        End If
        mHoverIdx = SendMessage(hWnd, LB_ITEMFROMPOINT, 0, ByVal Rct.Top * 65536)
        If MultiSelect Then SendMessage hWnd, LB_SETCURSEL, mHoverIdx, ByVal 0&
        If mHoverIdx > -1 And mHoverIdx < ListCount Then RaiseEvent MouseMoveOnItem(Rct.Left, Rct.Top Mod mItemHeight)
        DefCall = False
    End Select
    
    If Closing And Not DroppedDownState Then
       Closing = False: RaiseEvent RollUp
       Me.Refresh
    End If
  If Err Then Err.Clear
End Sub
Private Sub oSC_WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, Ret As Long, DefCall As Boolean)
  On Error Resume Next
    Dim Drw As DRAWITEMSTRUCT, SI As SCROLLINFO, Rct As RECT, Offs&(0 To 1)
    Const WM_DRAWITEM = &H2B
    
    If Msg = WM_DRAWITEM Then
       CopyMemory Drw, ByVal lParam, Len(Drw)
       GetWindowOrgEx Drw.hDC, Offs(0)
       GetClientRect hWndLB, Rct
          SI.cbSize = LenB(SI): SI.fMask = 31
          SI.nMax = HorizontalExtent
          SI.nPage = Rct.Right
          SI.nPos = Offs(0)
       SetScrollInfo hWndLB, 0, SI, 1
       ShowScrollBar hWndLB, 0, IIf(HorizontalExtent > Rct.Right, 1, 0)
 
       If oBB.ScaleWidth < Drw.rcItem.Right - Drw.rcItem.Left Or oBB.ScaleHeight < Drw.rcItem.Bottom - Drw.rcItem.Top Then
          oBB.Move 0, 0, ScaleX(Drw.rcItem.Right - Drw.rcItem.Left, vbPixels, ScaleMode), ScaleY(Drw.rcItem.Bottom - Drw.rcItem.Top, vbPixels, ScaleMode)
       End If
       oBB.Cls
       RaiseEvent OwnerDraw(Drw.itemID, CBool(Drw.itemState And 1), Drw.rcItem.Left > 0, oBB, Drw.rcItem.Right - Drw.rcItem.Left, Drw.rcItem.Bottom - Drw.rcItem.Top)
       BitBlt Drw.hDC, Drw.rcItem.Left, Drw.rcItem.Top, Drw.rcItem.Right - Drw.rcItem.Left, Drw.rcItem.Bottom - Drw.rcItem.Top, oBB.hDC, 0, 0, vbSrcCopy
       DefCall = False: Ret = 1
    End If
  If Err Then Err.Clear
End Sub
 
'***** Public Interface-Methods
Public Property Get ItemHeight() As Long
  ItemHeight = mItemHeight
End Property
Public Property Let ItemHeight(ByVal RHS As Long)
  mItemHeight = IIf(RHS < 4, 4, RHS)
  SendMessage oCB.hWnd, &H153, -1, ByVal mItemHeight
  SendMessage oCB.hWnd, &H153, 0, ByVal mItemHeight
  UserControl_Resize
End Property

Public Property Get MinVisibleItems() As Long
  MinVisibleItems = mMinVisible
End Property
Public Property Let MinVisibleItems(ByVal RHS As Long)
  Const CB_SETMINVISIBLE = &H1701
  mMinVisible = IIf(RHS < 4, 4, RHS)
  SendMessage oCB.hWnd, CB_SETMINVISIBLE, mMinVisible, ByVal 0&
  Dim Rct As RECT: GetClientRect oCB.hWnd, Rct
  MoveWindow oCB.hWnd, 0, 0, Rct.Right, mItemHeight * (mMinVisible + 1) + 8, 0
End Property

Public Property Get ListCount() As Long
  ListCount = oCB.ListCount
End Property
Public Property Let ListCount(ByVal RHS As Long)
  Const LB_ADDSTRING = &H180, LB_DELETESTRING = &H182, CB_RESETCONTENT = &H14B, S$ = ""
  If RHS < oCB.ListCount \ 2 Then SendMessage oCB.hWnd, CB_RESETCONTENT, 0, 0&
  Dim Msg As Long, i As Long
      Msg = IIf(RHS - oCB.ListCount > 0, LB_ADDSTRING, LB_DELETESTRING)
  For i = 1 To Abs(RHS - oCB.ListCount): SendMessage hWndLB, Msg, 0, ByVal 0&: Next
End Property

Public Property Get DroppedDownState() As Boolean
  DroppedDownState = IsWindowVisible(hWndLB) <> 0
End Property

Public Property Get ListIndex() As Long
  ListIndex = oCB.ListIndex
End Property
Public Property Let ListIndex(ByVal RHS As Long)
  oCB.ListIndex = RHS
End Property

Public Property Get HoverIndex() As Long
  HoverIndex = mHoverIdx
End Property

Public Sub TextOut(x, y, ByVal S As String)
  TextOutW oBB.hDC, x, y, StrPtr(S), Len(S)
End Sub

Public Sub Refresh()
  If DroppedDownState Then RedrawWindow hWndLB, 0, 0, &H101&
  RedrawWindow oCB.hWnd, 0, 0, &H101&
End Sub

Public Property Get HorizontalExtent() As Long
  Const LB_GETHORIZONTALEXTENT = &H193
  HorizontalExtent = SendMessage(hWndLB, LB_GETHORIZONTALEXTENT, 0, ByVal 0&)
End Property
Public Property Let HorizontalExtent(ByVal RHS As Long)
  Const LB_SETHORIZONTALEXTENT = &H194
  SendMessage hWndLB, LB_SETHORIZONTALEXTENT, RHS, ByVal 0&
End Property

