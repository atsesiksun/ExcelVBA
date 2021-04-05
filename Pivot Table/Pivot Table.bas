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
Sub PivotTable()
Dim a As Long
Dim b As Long
Dim c As Long
Dim d As Long
Dim x As Long
Dim y As Long

Sheets("Result1").UsedRange.Clear

'Create FullName Column
Sheets("Data").Columns(6).Insert
x = Sheets("Data").Range("A65000").End(xlUp).Row
Cells(1, 6) = "FullName"
For a = 2 To x
    Cells(a, 6) = Cells(a, 3) & " " & Cells(a, 5)
Next a

'Find distinct countries name and place them in Result1
Sheets("Data").Columns(19).Copy Sheets("Result1").Columns(1) ' Copy the countries column to result1
Sheets("Result1").Columns(1).RemoveDuplicates (1) 'Remove duplicates of the countries
With Sheets("Result1").Cells(1, 1)
    .Cells.Sort Key1:=.Columns(1), Order1:=xlAscending, Orientation:=xlTopToBottom, Header:=xlYes 'Sort the countries names
End With
y = Sheets("Result1").Range("A65000").End(xlUp).Row
For b = 2 To y
    Sheets("Result1").Cells(1, b) = Sheets("Result1").Cells(b, 1) 'rearrange the countries to rows
Next b
Sheets("Result1").Cells(1, 1) = "Customer Name"

'Copy the names and sales amount to Result1 according to the countries and if name is not red
c = 2
For a = 2 To x
    For b = 2 To y
        If Sheets("Data").Cells(a, 19) = Sheets("Result1").Cells(1, b) And Sheets("Data").Cells(a, 3).Font.Color <> vbRed Then
            Sheets("Result1").Cells(c, 1) = Sheets("Data").Cells(a, 6)
            Sheets("Result1").Cells(c, b) = Sheets("Data").Cells(a, 7)
            c = c + 1
        End If
    Next b
Next a

'Sort the names in Result1
With Sheets("Result1").Cells(1, 1)
    .Cells.Sort Key1:=.Columns(1), Order1:=xlAscending, Orientation:=xlTopToBottom, Header:=xlYes
End With

'Remove duplicates of names in result1 and add the sales amount together
x = Sheets("Result1").Range("A65000").End(xlUp).Row
For a = 2 To x
    If a <= x Then
        If Sheets("result1").Cells(a, 1) = Sheets("result1").Cells(a - 1, 1) Then
            For b = 2 To y
                If IsEmpty(Sheets("result1").Cells(a, b)) = False Then
                    Sheets("result1").Cells(a - 1, b) = Sheets("result1").Cells(a, b).Value + Sheets("result1").Cells(a - 1, b).Value
                End If
            Next b
            Sheets("result1").Cells(a, b).EntireRow.Delete
            a = a - 1
            x = x - 1
        End If
    End If
Next a
    
'Calculate the grand totals
x = Sheets("Result1").Range("A65000").End(xlUp).Row
Sheets("Result1").Cells(x + 1, 1) = "Grand Total"
For b = 2 To y
c = 0
    For a = 2 To x
        d = Sheets("result1").Cells(a, b).Value
        c = c + d
    Next a
    Sheets("Result1").Cells(x + 1, b) = c
Next b

End Sub
