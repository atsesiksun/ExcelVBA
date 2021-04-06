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
Sub ReadWriteTextFile()
Dim FilePath1 As String
Dim FilePath2 As String
Dim TextFile As Integer
Dim arr(1 To 10) As String
Dim i As Integer
Dim j As Integer

FilePath1 = "C:\Users\....\FirstLastNames1.txt"
FilePath2 = "C:\Users\....\FirstLastNames2.txt"

TextFile = FreeFile

Open FilePath1 For Input As TextFile 'Open Input File

'Access each line in text file and store them in array
i = 1
Do Until EOF(1) 'loop until end of file
    Line Input #TextFile, Textline 'Read line into variable
    arr(i) = Textline
    i = i + 1
Loop
Close TextFile

'Rearrange names to ascending order
For i = 1 To 10
    For j = i + 1 To 10
        a = arr(i)
        If Len(arr(i)) > Len(arr(j)) Then
            arr(i) = arr(j)
            arr(j) = a
        End If
    Next j
Next i

'Create ouput file and store results in it
Open FilePath2 For Output As TextFile
For i = 1 To 10
    Print #TextFile, arr(i)
Next i
Close TextFile

End Sub
