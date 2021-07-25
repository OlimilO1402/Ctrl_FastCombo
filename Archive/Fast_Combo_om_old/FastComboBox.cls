VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FastComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long

Private Const BOOL_FALSE     As Long = 0
Private Const BOOL_TRUE      As Long = 1
Private Const WM_SETREDRAW   As Long = &HB

Private Const CB_ADDSTRING   As Long = &H143
Private Const CB_INITSTORAGE As Long = &H161

'Private WithEvents mCombo As ComboBox
Private mCombo As ComboBox

Dim sl() As String
Dim TestResuls As New Collection

Friend Sub New_(Combo As ComboBox)
    Set mCombo = Combo
    'Debug.Print mCombo.hWnd
End Sub

Public Property Get hWnd() As Long
    hWnd = mCombo.hWnd
End Property

'Public Function FastComboBox(aCombo As ComboBox) As FastComboBox
'    Set FastComboBox = New FastComboBox: FastComboBox.New_ aCombo
'End Function

Private Function StrArr_WToA(sl() As String) As String()
    Dim i As Long, u As Long: u = UBound(sl)
    ReDim asl(0 To u) As String
    For i = 0 To u
        asl(i) = StrConv(sl(i), vbFromUnicode) 'Narrow)
    Next
    StrArr_WToA = asl
End Function

Public Property Let List(RHS() As String)
    With mCombo
        
        Dim sl() As String: sl = StrArr_WToA(RHS)
        
        Dim CBhWnd As Long: CBhWnd = .hWnd
        
        Dim i As Long
        Dim u As Long: u = UBound(sl)
        Dim n As Long: n = u + 1
        Dim hr As Long
        
        hr = SendMessage(CBhWnd, WM_SETREDRAW, ByVal BOOL_FALSE, ByVal 0&)
        
        hr = SendMessage(CBhWnd, CB_INITSTORAGE, ByVal n, ByVal 20 * n)
        
        Dim lsl() As Long: SAPtr(ArrPtr(lsl)) = SAPtr(StrArrPtr(sl))
        For i = 0 To u
            
            hr = SendMessage(CBhWnd, CB_ADDSTRING, ByVal 0&, ByVal lsl(i))
            
        Next
        ZeroSAPtr ArrPtr(lsl)
        
        hr = SendMessage(CBhWnd, WM_SETREDRAW, ByVal BOOL_TRUE, ByVal 0&)
        
        .Refresh
    End With
End Property
Public Property Get List() As String()
    With mCombo
        
        Dim i As Long, u As Long: u = .ListCount - 1
        ReDim sl(0 To u) As String
        
        For i = 0 To u
            sl(i) = .List(i)
        Next
        
    End With
    List = sl
End Property

