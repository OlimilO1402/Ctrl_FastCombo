Attribute VB_Name = "MComboBox"
Option Explicit
'Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long

Private Const BOOL_FALSE     As Long = 0
Private Const BOOL_TRUE      As Long = 1
Private Const WM_SETREDRAW   As Long = &HB&

Private Const CB_ADDSTRING   As Long = &H143&
Private Const CB_INITSTORAGE As Long = &H161&

Private Const LBS_SORT           As Long = &H2&
Private Const LBS_NOREDRAW       As Long = &H4&
Private Const LBS_OWNERDRAWFIXED As Long = &H10&
Private Const LBS_HASSTRINGS     As Long = &H40&
Private Const LBS_NODATA         As Long = &H2000&

Private Const LB_SETCOUNT        As Long = &H1A7&


Private Const CBS_OWNERDRAWFIXED As Long = &H10&
Private Const CBS_SORT           As Long = &H100&
Private Const CBS_HASSTRINGS     As Long = &H200&



Dim sl() As String
Dim TestResults As New Collection

Public Property Let ComboBox_List(this As ComboBox, sl() As String)
    Dim i As Long
    With this

        Dim CBhWnd As Long: CBhWnd = .hWnd

        Dim u As Long: u = UBound(sl)
        Dim n As Long: n = u + 1
        Dim hr As Long

'LBS_NOREDRAW
        'hr = SendMessage(CBhWnd, LBS_NOREDRAW, ByVal BOOL_FALSE, ByVal 0&)

'LBS_NODATA
        'hr = SendMessage(CBhWnd, LBS_NODATA, ByVal BOOL_FALSE, ByVal 0&)

'LB_SETCOUNT
        hr = SendMessage(CBhWnd, LB_SETCOUNT, ByVal n, ByVal 0&)

'CBS_OWNERDRAWFIXED
        'hr = SendMessage(CBhWnd, CBS_OWNERDRAWFIXED, ByVal BOOL_FALSE, ByVal 0&)

'CBS_SORT = False
'CBS_HASSTRINGS = False
        'hr = SendMessage(CBhWnd, CBS_SORT, ByVal BOOL_FALSE, ByVal 0&)
        'hr = SendMessage(CBhWnd, CBS_HASSTRINGS, ByVal BOOL_FALSE, ByVal 0&)

        hr = SendMessage(CBhWnd, WM_SETREDRAW, ByVal BOOL_FALSE, ByVal 0&)

        hr = SendMessage(CBhWnd, CB_INITSTORAGE, ByVal n, ByVal 20 * n)

        'Dim lsl() As Long: SAPtr(ArrPtr(lsl)) = SAPtr(StrArrPtr(sl))
        For i = 0 To u
            'hr = SendMessage(CBhWnd, CB_ADDSTRING, ByVal 0&, ByVal lsl(i))
            hr = SendMessage(CBhWnd, CB_ADDSTRING, ByVal 0&, ByVal StrPtr(sl(i)))
        Next
        'ZeroSAPtr ArrPtr(lsl)
        'hr = SendMessage(CBhWnd, CBS_OWNERDRAWFIXED, ByVal BOOL_TRUE, ByVal 0&)
        '.Visible = True
        hr = SendMessage(CBhWnd, WM_SETREDRAW, ByVal BOOL_TRUE, ByVal 0&)
        
        .Refresh
    End With
End Property
Public Property Get ComboBox_List(this As ComboBox) As String()
    With this
        Dim u As Long: u = .ListCount - 1
        ReDim sl(0 To u) As String
        Dim i As Long
        For i = 0 To u
            sl(i) = .List(i)
        Next
    End With
    ComboBox_List = sl
End Property
