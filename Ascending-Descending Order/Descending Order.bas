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
Sub DescendingOrder()
Dim arr(0, 6) As Integer
Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim y As Integer

For i = 0 To 6
    arr(0, i) = Sheet1.Cells(1, i + 1)
    y = i
    If i > 0 Then
            x = arr(0, y)
            For j = i - 1 To 0 Step -1
                If arr(0, y) > arr(0, j) Then
                    arr(0, y) = arr(0, y - 1)
                    Sheet1.Cells(2, y + 1) = arr(0, y)
                    arr(0, j) = x
                    Sheet1.Cells(2, j + 1) = arr(0, j)
                    y = y - 1
                Else
                    arr(0, i) = arr(0, i)
                    Sheet1.Cells(2, i + 1) = arr(0, i)
                End If
            Next j
    Else
    Sheet1.Cells(2, i + 1) = arr(0, i)
    End If
Next i


End Sub
