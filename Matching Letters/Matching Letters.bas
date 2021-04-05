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
Private Sub MatchingLetters()
Dim arr1(1 To 10) As String
Dim arr2(1 To 26) As String
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim i As Integer
Dim j As Integer
Dim m As Integer
Dim n As Integer
Dim x As Integer
Dim y As String
Dim z As Integer
Dim w As String

arr1(1) = "Jason"
arr1(2) = "Kimberley"
arr1(3) = "Daryl"
arr1(4) = "Brandon"
arr1(5) = "Michelle"
arr1(6) = "Emilie"
arr1(7) = "Christine"
arr1(8) = "Nathalie"
arr1(9) = "Joon"
arr1(10) = "Andrew"

Debug.Print "Names:"
For i = 1 To 10
    Debug.Print (arr1(i)) 'Print the 10 human names
Next i

Debug.Print " "
Debug.Print "Matching Letters:"
b = 1   ' base to track arr2
For i = 1 To 10
    x = Len(arr1(i)) 'track the length of first name
    For j = 1 To x
        y = LCase(Mid(arr1(i), j, 1)) ' track each letter in a name and convert uppercase to lowercase
        For m = i + 1 To 10
            z = Len(arr1(m)) 'track the length of second name
            For n = 1 To z
                w = LCase(Mid(arr1(m), n, 1)) ' track each letter in a name and convert uppercase to lowercase
                a = 0
                For c = 1 To 26
                    If y = arr2(c) Then ' track number of y in arr2
                    a = a + 1
                    End If
                Next c
                If y = w And a = 0 Then ' if matching and not existing in arr2, then insert the letter in arr2
                    arr2(b) = y
                    Debug.Print arr2(b)
                    b = b + 1
                End If
            Next n
        Next m
    Next j
Next i

End Sub
