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
Sub NameVlookup()

Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim x As Integer
Dim y As Integer
Dim z As Integer
Dim SearchStr As String
Dim StrFound As Range

SearchStr = "~*"
Set StrFound = Sheets("Sheet2").Range("A:Z").Find(SearchStr)
a = StrFound.Column 'column position of *
b = StrFound.Row 'row position of *
c = Sheet1.Range("A65000").End(xlUp).Row 'last used row in column A in sheet1

'Clear column B values before doing lookup
Sheet1.Range("B2:B" & c).Clear

For x = 2 To c ' track each ID in sheet 1
    For y = 1 To a 'track each column until * in sheet 2
        For z = 1 To b 'track each colum until * in sheet 2
            If Sheet1.Cells(x, 1) = Sheet2.Cells(z, y) And y <> a Then 'if id in sheet1 equals id in sheet2 and if id is not in the same column as *
                Sheet1.Cells(x, 2) = Sheet2.Cells(z, y + 1)
            End If
        Next z
    Next y
Next x
   
End Sub


