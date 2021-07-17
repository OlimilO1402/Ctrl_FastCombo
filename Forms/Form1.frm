VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3975
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   960
      Width           =   3975
   End
   Begin VB.CommandButton BtnTestFillCombo 
      Caption         =   "Fill Combo in Form"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long

Private Const BOOL_FALSE     As Long = 0
Private Const BOOL_TRUE      As Long = 1
Private Const WM_SETREDRAW   As Long = &HB

Private Const CB_ADDSTRING   As Long = &H143
Private Const CB_INITSTORAGE As Long = &H161


Dim sl() As String
Dim TestResuls As New Collection

Private Sub BtnTestFillCombo_Click()
    
    Dim dt As Single
    
    dt = Timer
    sl = GetStrArr(100000)
    dt = Timer - dt
    
    TestResuls.Add "Fill: sl(0) = " & sl(0)
    TestResuls.Add "time dt = " & dt & " s"
    Text1.Text = GetTestResults
    DoEvents
    
    dt = Timer
    sl = StrArr_WToA(sl)
    dt = Timer - dt
    
    TestResuls.Add "Ansi convert finished"
    TestResuls.Add "time dt = " & dt & " s"
    Text1.Text = GetTestResults
    DoEvents
    Combo1.Visible = False
    DoEvents
    dt = Timer
    Me.ComboBox_List(Combo1) = sl
    dt = Timer - dt
    
    Combo1.Visible = True
    
    TestResuls.Add "Combo fill finished"
    TestResuls.Add "time dt = " & dt & " s"
    Text1.Text = GetTestResults
    
End Sub

Function GetTestResults() As String
    Dim s As String
    Dim i As Long
    For i = 1 To TestResuls.Count
        s = s & TestResuls(i) & vbCrLf
    Next
    GetTestResults = s
End Function

Function GetStrArr(ByVal n As Long) As String()
    Dim i As Long, dx As Double: dx = 1 / n
    ReDim sl(0 To n) 'As String
    For i = 0 To n
        sl(i) = Format((n - i) * dx, "0.00000") & vbNullChar
    Next
    GetStrArr = sl
End Function

Private Function StrArr_WToA(sl() As String) As String()
    Dim i As Long, u As Long: u = UBound(sl)
    ReDim asl(0 To u) As String
    For i = 0 To u
        asl(i) = StrConv(sl(i), vbFromUnicode) 'Narrow)
    Next
    StrArr_WToA = asl
End Function

'' #################### ' Modul MCombo ' #################### '
Public Property Let ComboBox_List(this As ComboBox, sl() As String)
    Dim i As Long
    With this

        Dim CBhWnd As Long: CBhWnd = .hWnd

        Dim u As Long: u = UBound(sl)
        Dim n As Long: n = u + 1
        Dim hr As Long

        hr = SendMessage(CBhWnd, WM_SETREDRAW, ByVal BOOL_FALSE, ByVal 0&)

        hr = SendMessage(CBhWnd, CB_INITSTORAGE, ByVal n, ByVal 20 * n)
        
        Dim lsl() As Long: SAPtr(ArrPtr(lsl)) = SAPtr(StrArrPtr(sl))
        For i = 0 To u
            hr = SendMessage(CBhWnd, CB_ADDSTRING, ByVal CLng(0), ByVal lsl(i))
        Next
        ZeroSAPtr ArrPtr(lsl)

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

Private Sub Form_Resize()
    Dim L As Single, T As Single, W As Single, H As Single
    L = 0: T = Combo1.Top
    W = Me.ScaleWidth: H = Combo1.Height
    If W > 0 And H > 0 Then Combo1.Move L, T, W ', H
    T = BtnTestFillCombo.Top
    H = BtnTestFillCombo.Height
    If W > 0 And H > 0 Then BtnTestFillCombo.Move L, T, W, H
    T = Text1.Top
    H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub
