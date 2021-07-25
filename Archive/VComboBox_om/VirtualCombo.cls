VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType    As Long
    CtlID      As Long
    itemID     As Long
    itemAction As Long
    itemState  As Long
    hwndItem   As Long
    hDC        As Long
    rcItem     As RECT
    ItemData   As Long
End Type

Private Type SCROLLINFO
    cbSize    As Long
    fMask     As Long
    nMin      As Long
    nMax      As Long
    nPage     As Long
    nPos      As Long
    nTrackPos As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (Dst As Any, Src As Any, ByVal Length As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TextOutW Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Declare Function GetWindowOrgEx Lib "gdi32" (ByVal hDC As Long, lpPoint As Any) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Any) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, ByVal lprcUpdate As Long, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal enmBar As Long, ByVal bShow As Long) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal enmBar As Long, lpSI As SCROLLINFO) As Long
Private Declare Function SetScrollInfo Lib "user32" (ByVal hWnd As Long, ByVal enmBar As Long, lpSI As SCROLLINFO, ByVal bRedraw As Long) As Long

Private WithEvents Owner As PictureBox
Attribute Owner.VB_VarHelpID = -1

Private WithEvents mCB As VB.ComboBox
Attribute mCB.VB_VarHelpID = -1
Private mPB As VB.PictureBox
Attribute mPB.VB_VarHelpID = -1
Private WithEvents oSL As cSubClass
Attribute oSL.VB_VarHelpID = -1
Private WithEvents oSC As cSubClass
Attribute oSC.VB_VarHelpID = -1
Private mItemHeight As Long
Private mMinVisible As Long
Private hWndLB As Long
Private mHoverIdx As Long

Public MultiSelect As Boolean
Public MouseWheelItemScrollDistance As Long

Public Event Click()
Public Event RollUp()
Public Event DropDown()
Public Event ListMultiClick()
Public Event KeyPress(KeyAscii As Integer, sKey As String)
Public Event MouseMoveOnItem(ByVal x As Long, ByVal y As Long)
Public Event OwnerDraw(ByVal Index As Long, ByVal IsSelected As Boolean, ByVal IsComboItem As Boolean, Canvas As PictureBox, ByVal dx As Long, ByVal dy As Long)

'Private Sub Class_Initialize()
'    Set mPB = Controls.Add("VB.PictureBox", "mPB")
'        mPB.AutoRedraw = True
'        mPB.BorderStyle = 0
'        mPB.ScaleMode = vbPixels
'        mPB.BackColor = &H80000005
'    hCreateHook = SetWindowsHookEx(WH_CBT, AddressOf modCreateHook.CreateHookProcForVBCombo, 0, App.ThreadID)
'    Set mCB = Controls.Add("VB.ComboBox", "mCB")
'    UnhookWindowsHookEx hCreateHook
'    hWndLB = hComboLBox
'    mCB.Visible = True
'    ItemHeight = 19
'    MinVisibleItems = 10
'    MouseWheelItemScrollDistance = 3
'End Sub

Friend Sub New_(aParent As Form, aOwner As PictureBox)
    Set Owner = aOwner
    Set mPB = Parent.Controls.Add("VB.PictureBox", "mPB")
    mPB.AutoRedraw = True
    mPB.BorderStyle = 0
    mPB.ScaleMode = vbPixels
    mPB.BackColor = &H80000005
    hCreateHook = SetWindowsHookEx(WH_CBT, AddressOf modCreateHook.CreateHookProcForVBCombo, 0, App.ThreadID)
    Set mCB = Controls.Add("VB.ComboBox", "mCB")
    UnhookWindowsHookEx hCreateHook
    hWndLB = hComboLBox
    mCB.Visible = True
    ItemHeight = 19
    MinVisibleItems = 10
    MouseWheelItemScrollDistance = 3
End Sub

Private Sub Owner_Paint()
    If Not oSC Is Nothing Then oSC.UnHook
    If Not oSL Is Nothing Then oSL.UnHook
    If Not Ambient_UserMode Then Exit Sub
    Set oSC = New cSubClass: oSC.Hook Owner.hWnd
    Set oSL = New cSubClass: oSL.Hook hWndLB
End Sub
Private Sub Owner_Hide()
    If Not oSC Is Nothing Then oSC.UnHook
    If Not oSL Is Nothing Then oSL.UnHook
End Sub
Private Sub Owner_Resize()
    On Error Resume Next
    mCB.Move 0, 0, Owner.Width
    Owner.Height = mCB.Height
End Sub
'Private Sub Owner_EnterFocus()
Private Sub Owner_GotFocus()
    mCB.SetFocus
End Sub
Private Sub Owner_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii, ChrW$(KeyAscii))
  If DroppedDownState And KeyAscii = vbKeySpace And ListIndex > -1 Then
    RaiseEvent ListMultiClick
    Me.Refresh
  End If
End Sub

'***** VB.Combo-EventHandlers
Private Sub mCB_Click()
    RaiseEvent Click
    Owner.Refresh
End Sub
Private Sub mCB_DropDown()
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
        If mCB.TopIndex - Sgn(wParam And &HFFFF) * MouseWheelItemScrollDistance < 0 Then
          mCB.TopIndex = 0
        Else
          mCB.TopIndex = mCB.TopIndex - Sgn(wParam And &HFFFF) * MouseWheelItemScrollDistance
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
 
       If mPB.ScaleWidth < Drw.rcItem.Right - Drw.rcItem.Left Or mPB.ScaleHeight < Drw.rcItem.Bottom - Drw.rcItem.Top Then
          mPB.Move 0, 0, ScaleX(Drw.rcItem.Right - Drw.rcItem.Left, vbPixels, ScaleMode), ScaleY(Drw.rcItem.Bottom - Drw.rcItem.Top, vbPixels, ScaleMode)
       End If
       mPB.Cls
       RaiseEvent OwnerDraw(Drw.itemID, CBool(Drw.itemState And 1), Drw.rcItem.Left > 0, mPB, Drw.rcItem.Right - Drw.rcItem.Left, Drw.rcItem.Bottom - Drw.rcItem.Top)
       BitBlt Drw.hDC, Drw.rcItem.Left, Drw.rcItem.Top, Drw.rcItem.Right - Drw.rcItem.Left, Drw.rcItem.Bottom - Drw.rcItem.Top, mPB.hDC, 0, 0, vbSrcCopy
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
    SendMessage mCB.hWnd, &H153, -1, ByVal mItemHeight
    SendMessage mCB.hWnd, &H153, 0, ByVal mItemHeight
    Owner_Resize
End Property

Public Property Get MinVisibleItems() As Long
    MinVisibleItems = mMinVisible
End Property
Public Property Let MinVisibleItems(ByVal RHS As Long)
    Const CB_SETMINVISIBLE = &H1701
    mMinVisible = IIf(RHS < 4, 4, RHS)
    SendMessage mCB.hWnd, CB_SETMINVISIBLE, mMinVisible, ByVal 0&
    Dim Rct As RECT: GetClientRect mCB.hWnd, Rct
    MoveWindow mCB.hWnd, 0, 0, Rct.Right, mItemHeight * (mMinVisible + 1) + 8, 0
End Property

Public Property Get ListCount() As Long
    ListCount = mCB.ListCount
End Property
Public Property Let ListCount(ByVal RHS As Long)
    Const LB_ADDSTRING = &H180, LB_DELETESTRING = &H182, CB_RESETCONTENT = &H14B, S$ = ""
    If RHS < mCB.ListCount \ 2 Then SendMessage mCB.hWnd, CB_RESETCONTENT, 0, 0&
    Dim Msg As Long, i As Long
    Msg = IIf(RHS - mCB.ListCount > 0, LB_ADDSTRING, LB_DELETESTRING)
    For i = 1 To Abs(RHS - mCB.ListCount): SendMessage hWndLB, Msg, 0, ByVal 0&: Next
End Property

Public Property Get DroppedDownState() As Boolean
    DroppedDownState = IsWindowVisible(hWndLB) <> 0
End Property

Public Property Get ListIndex() As Long
    ListIndex = mCB.ListIndex
End Property
Public Property Let ListIndex(ByVal RHS As Long)
    mCB.ListIndex = RHS
End Property

Public Property Get HoverIndex() As Long
    HoverIndex = mHoverIdx
End Property

Public Sub TextOut(x, y, ByVal S As String)
    TextOutW mPB.hDC, x, y, StrPtr(S), Len(S)
End Sub

Public Sub Refresh()
    If DroppedDownState Then RedrawWindow hWndLB, 0, 0, &H101&
    RedrawWindow mCB.hWnd, 0, 0, &H101&
End Sub

Public Property Get HorizontalExtent() As Long
    Const LB_GETHORIZONTALEXTENT = &H193
    HorizontalExtent = SendMessage(hWndLB, LB_GETHORIZONTALEXTENT, 0, ByVal 0&)
End Property
Public Property Let HorizontalExtent(ByVal RHS As Long)
    Const LB_SETHORIZONTALEXTENT = &H194
    SendMessage hWndLB, LB_SETHORIZONTALEXTENT, RHS, ByVal 0&
End Property
