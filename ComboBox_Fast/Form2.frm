VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3015
   LinkTopic       =   "Form2"
   ScaleHeight     =   3255
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnTestFillCombo 
      Caption         =   "Fill Combo in WrapperClass"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2775
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
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
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
Dim TestResuls As New Collection

Private Sub Form_Load()
    Set FCombo = FastComboBox(Combo1)
End Sub

Private Sub BtnTestFillCombo_Click()
    
    Dim dt As Single
    
    dt = Timer
    sl = GetStrArr(100000)
    dt = Timer - dt
    
    TestResuls.Add "Fill: sl(0) = " & sl(0)
    TestResuls.Add "time dt = " & dt & " s"
    Text1.Text = GetTestResults
    DoEvents
    
'    dt = Timer
'    sl = StrArr_WToA(sl)
'    dt = Timer - dt
'
'    TestResuls.Add "Ansi convert finished"
'    TestResuls.Add "time dt = " & dt & " s"
'    Text1.Text = GetTestResults
'    DoEvents
    
    dt = Timer
    FCombo.List = sl
    dt = Timer - dt
    
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
