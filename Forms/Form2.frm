VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3975
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3255
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTestFillCombo 
      Caption         =   "Fill Combo in WrapperClass"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   480
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sl() As String
Dim FCombo As FastComboBox
Dim TestResults As New Collection

Private Sub Form_Load()
    Set FCombo = FastComboBox(Combo1)
End Sub

Private Sub BtnTestFillCombo_Click()
    
    Dim dt As Single
    
    dt = Timer
    sl = GetStrArr(100000)
    dt = Timer - dt
    
    TestResults.Add "Fill: sl(0) = " & sl(0)
    TestResults.Add "time dt = " & dt & " s"
    Text1.Text = GetTestResults
    DoEvents
    
'    dt = Timer
'    sl = StrArr_WToA(sl)
'    dt = Timer - dt
'
'    TestResults.Add "Ansi convert finished"
'    TestResults.Add "time dt = " & dt & " s"
'    Text1.Text = GetTestResults
'    DoEvents
    
    dt = Timer
    FCombo.List = sl
    dt = Timer - dt
    
    TestResults.Add "Combo fill finished"
    TestResults.Add "time dt = " & dt & " s"
    Text1.Text = GetTestResults
    
End Sub

Function GetTestResults() As String
    Dim s As String
    Dim i As Long
    For i = 1 To TestResults.Count
        s = s & TestResults(i) & vbCrLf
    Next
    GetTestResults = s
End Function

Function GetStrArr(ByVal n As Long) As String()
    Dim i As Long, dx As Double: dx = 1 / n
    ReDim sl(0 To n) 'As String
    For i = 0 To n
        sl(i) = Format((n - i) * dx, "0.00000")
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

