Attribute VB_Name = "MArray"
Option Explicit
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dst As Any, ByRef Src As Any, ByVal BytLength As Long)
Private Declare Sub RtlZeroMemory Lib "kernel32" (ByRef Dst As Any, ByVal BytLength As Long)
'die Funktion ArrPtr geht bei allen Arrays außer bei String-Arrays
Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef Arr() As Any) As Long
'Private Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (ByRef Arr() As Any) As Long


'deswegen hier eine Hilfsfunktion für StringArrays
Public Function StrArrPtr(ByRef strArr As Variant) As Long
    Call RtlMoveMemory(StrArrPtr, ByVal VarPtr(strArr) + 8, 4)
End Function
'e.g. how to use it:
'    SAPtr(StrArrPtr(strArr2)) = SAPtr(StrArrPtr(strArr1))

'jetzt kann das Property SAPtr für Alle Arrays verwendet werden,
'um den Zeiger auf den Safe-Array-Descriptor eines Arrays einem
'anderen Array zuzuweisen.
Public Property Get SAPtr(ByVal pArr As Long) As Long
    Call RtlMoveMemory(SAPtr, ByVal pArr, 4)
End Property

Public Property Let SAPtr(ByVal pArr As Long, ByVal RHS As Long)
    Call RtlMoveMemory(ByVal pArr, RHS, 4)
End Property

Public Sub ZeroSAPtr(ByVal pArr As Long)
    Call RtlZeroMemory(ByVal pArr, 4)
End Sub



