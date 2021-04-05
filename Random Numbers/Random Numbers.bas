VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub RandomNumbers()
Dim arr(2, 1) As Integer
Dim i As Integer
Dim j As Integer

Sheet1.UsedRange.Clear

For i = 0 To 2
    For j = 0 To 1
    arr(i, j) = Int(50 * Rnd) + 1
    Sheet1.Cells(i + 1, j + 1) = arr(i, j)
    Next j
Next i
End Sub

