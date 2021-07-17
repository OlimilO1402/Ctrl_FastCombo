Attribute VB_Name = "MMain"
Option Explicit

Sub Main()
    Form1.Show
    Form2.Show
    Form3.Show
End Sub

Public Function FastComboBox(aCombo As ComboBox) As FastComboBox
    Set FastComboBox = New FastComboBox: FastComboBox.New_ aCombo
End Function

