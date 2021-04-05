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
Sub IncrementalNumbers()
Dim arr(0, 4) As Integer
Dim i As Integer
Dim x As Integer

x = 0
For i = 0 To 4
    arr(0, i) = x + 1000
    Sheet1.Cells(1, i + 1).Value = arr(0, i)
    x = x + 1000
Next i

End Sub

