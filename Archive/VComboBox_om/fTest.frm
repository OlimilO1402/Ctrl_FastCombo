VERSION 5.00
Begin VB.Form fTest 
   Caption         =   "Virtual (OwnerDrawn) ComboBox"
   ClientHeight    =   4350
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   290
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox cboVSimple 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown-Liste
      TabIndex        =   2
      Top             =   600
      Width           =   5655
   End
   Begin VB.PictureBox PnlVCombo 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   5595
      TabIndex        =   1
      Top             =   1320
      Width           =   5655
   End
   Begin PVirtualCombo.ucVirtualCombo ucVCombo1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   661
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tCountries
    Nam As String
    Chk As Boolean
End Type
Private ResPath As String
Private CLst() As tCountries '<- this external DataSource-Container could also be a Recordset (e.g. to support easy Sorting)

Private WithEvents VComboBox1 As VComboBox
Attribute VComboBox1.VB_VarHelpID = -1


Private Sub Form_Load()
    ResPath = App.Path & "\Res\"
    'prepare an example for an external Data-Set (here in a UDT-Arr)
    Dim i As Long, F As String
    F = Dir(ResPath & "*.gif")
    Do While Len(F)
        ReDim Preserve CLst(i): CLst(i).Nam = Left(F, Len(F) - 4): i = i + 1
        F = Dir
    Loop
    
    'now setup the "MultiSelect-Combo"
    ucVCombo1.ListCount = UBound(CLst) + 1
    ucVCombo1.ItemHeight = 22
    ucVCombo1.MinVisibleItems = 15
    ucVCombo1.MultiSelect = True
    '  ucVCombo1.HorizontalExtent = 400
    
    'finally the setup for the "Simple-Combo"
    cboVSimple.ListCount = 32 '<- setting a ListCount is all what's needed here
    
    
    Set VComboBox1 = New VComboBox1
    VComboBox1.New_ Me, PnlVCombo
    
    
End Sub
 
'small helper, to join the currently selected Countries
Private Function GetCheckedCountries(Optional Delimiter As String = ", ") As String
    Dim i As Long, j As Long, Arr() As String
    For i = 0 To UBound(CLst)
        If CLst(i).Chk Then ReDim Preserve Arr(j): Arr(j) = CLst(i).Nam: j = j + 1
    Next
    GetCheckedCountries = Join(Arr, Delimiter)
End Function
 
'***** 4 Event-Handlers for the "Multi-Flag-Combo"-Scenario
Private Sub ucVCombo1_KeyPress(KeyAscii As Integer, sKey As String) 'a simple search-routine on cboV's external Country-Data
    Dim i As Long
    For i = 0 To UBound(CLst)
        If StrComp(Left$(CLst(i).Nam, 1), sKey, 1) = 0 Then ucVCombo1.ListIndex = i: Exit For
    Next
End Sub
Private Sub ucVCombo1_ListMultiClick()
    CLst(ucVCombo1.ListIndex).Chk = Not CLst(ucVCombo1.ListIndex).Chk 'toggle the Checked-State in our external Data
End Sub
Private Sub ucVCombo1_MouseMoveOnItem(ByVal x As Long, ByVal y As Long)
    'Debug.Print "MouseMoveOnItem (HoverIndex=" & ucVCombo1.HoverIndex & ", x=" & x & ", y=" & y & ")"
    ucVCombo1.Tag = IIf(x < 20, "x", ""): ucVCombo1.Refresh 'just to show, how a certain Hover-State within an Item can be handled (here we force a Hover-Effect over the CheckBox-area)
End Sub
Private Sub ucVCombo1_OwnerDraw(ByVal Index As Long, ByVal IsSelected As Boolean, ByVal IsComboItem As Boolean, Canvas As PictureBox, ByVal dx As Long, ByVal dy As Long)
    With Canvas 'all Drawing-Output happens "Item-wise" on a Canvas-PicBox, which is passed from inside the Virtual-ComboControl
        .FontName = "Arial": .FontSize = 10
        Canvas.Line (0, 0)-(dx, dy), IIf(IsSelected, RGB(212, 232, 255), Canvas.BackColor), BF
        
        If Index = -1 Or IsComboItem Then 'here we choose, to draw the "checked accumulation" when Index= -1 comes in
            ucVCombo1.TextOut 32, 3, IIf(Len(GetCheckedCountries), GetCheckedCountries, "<Select multiple Countries>")
            .PaintPicture LoadPicture(ResPath & CLst(1).Nam & ".gif"), 0, 0, 20, 15
            .PaintPicture LoadPicture(ResPath & CLst(4).Nam & ".gif"), 3, 3, 20, 15
            .PaintPicture LoadPicture(ResPath & CLst(7).Nam & ".gif"), 6, 6, 20, 15
        Else 'it has to be an Item of the DropDown-List
            ucVCombo1.TextOut 58, 2, CLst(Index).Nam  'Print the Country-Name
            'draw an empty Rectangle for the CheckBox (respecting the Mouse-Hover-State, when over the CheckBox-area of a certain-Item)
            Canvas.Line (4, 4)-(dy - 5, dy - 5), IIf(Len(ucVCombo1.Tag) And Index = ucVCombo1.HoverIndex, vbRed, vbBlack), B
            'now Print the Checked-State into the just drawn Rectangle
            If CLst(Index).Chk Then .FontName = "WebDings": .FontSize = 13: ucVCombo1.TextOut 3, 0, "a"
            .PaintPicture LoadPicture(ResPath & CLst(Index).Nam & ".gif"), 23, 1 'draw the Flag
        End If
    End With
End Sub

'***** and finally the two Event-Handlers for the Simple-Combo-scenario (which has no external DataSource, but renders its ListIndexes instead)
Private Sub cboVSimple_Click()
  Debug.Print "cboVSimple_Click", cboVSimple.ListIndex
End Sub
Private Sub cboVSimple_OwnerDraw(ByVal Index As Long, ByVal IsSelected As Boolean, ByVal IsComboItem As Boolean, Canvas As PictureBox, ByVal dx As Long, ByVal dy As Long)
  Canvas.Line (0, 0)-(dx, dy), IIf(IsSelected, RGB(205, 230, 255), Canvas.BackColor), BF
  Canvas.FontName = "Arial": Canvas.FontSize = 10
  cboVSimple.TextOut 1, 1, Index
End Sub
