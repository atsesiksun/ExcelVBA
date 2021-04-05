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
Sub PostalCode()
Dim RndNum1 As Integer
Dim RndNum2 As Integer
Dim RndNum3 As Integer
Dim RndAlpha1 As String
Dim RndAlpha2 As String
Dim RndAlpha3 As String
Dim PostalCode2 As String
Dim x As Long
Dim a As Long

x = Sheets(1).Range("A65000").End(xlUp).Row

RndNum1 = Int((9 - 0 + 1) * Rnd + 0)
RndNum2 = Int((9 - 0 + 1) * Rnd + 0)
RndNum3 = Int((9 - 0 + 1) * Rnd + 0)
RndAlpha1 = Chr(Int((90 - 65 + 1) * Rnd + 65))
RndAlpha2 = Chr(Int((90 - 65 + 1) * Rnd + 65))
RndAlpha3 = Chr(Int((90 - 65 + 1) * Rnd + 65))

PostalCode2 = RndAlpha1 & CStr(RndNum1) & RndAlpha2 & " " & CStr(RndNum2) & RndAlpha3 & CStr(RndNum3)
Debug.Print PostalCode2

For a = 1 To x
    If Cells(a, 1).Value = PostalCode2 Then
        RndNum1 = Int((9 - 0 + 1) * Rnd + 0)
        RndNum2 = Int((9 - 0 + 1) * Rnd + 0)
        RndNum3 = Int((9 - 0 + 1) * Rnd + 0)
        RndAlpha1 = Chr(Int((90 - 65 + 1) * Rnd + 65))
        RndAlpha2 = Chr(Int((90 - 65 + 1) * Rnd + 65))
        RndAlpha3 = Chr(Int((90 - 65 + 1) * Rnd + 65))
        PostalCode2 = RndAlpha1 & CStr(RndNum1) & RndAlpha2 & " " & CStr(RndNum2) & RndAlpha3 & CStr(RndNum3)
    Else
        Cells(x + 1, 1).Value = PostalCode2
    End If
Next a

End Sub
