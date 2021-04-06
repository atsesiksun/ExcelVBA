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
Private Sub FirstLastName()
Dim arr1(1 To 10) As String
Dim i As Integer
Dim j As Integer
Dim a As Integer
Dim b As String
Dim c As String
Dim d As String
Dim e As Integer


arr1(1) = "KimberleyQiu"
arr1(2) = "Brandon Lcc"
arr1(3) = "PhoebeLiu"
arr1(4) = "Christopher Chan"
arr1(5) = "ChristinaShang"
arr1(6) = "Warren Wan"
arr1(7) = "MichelleChoi"
arr1(8) = "Joon Huh"
arr1(9) = "AndrewIp"
arr1(10) = "Daryl Lim"

'Print the 10 human names
Debug.Print "First-Last names:"
For i = 1 To 10
 Debug.Print arr1(i)
Next i

'Rearrange the names by length
For i = 1 To 10
b = arr1(i)
    For j = i + 1 To 10
        If Len(arr1(i)) > Len(arr1(j)) Then
            arr1(i) = arr1(j)
            arr1(j) = b
            b = arr1(i)
        End If
    Next j
Next i

'Change the names to last and first name
Debug.Print " "
Debug.Print "Last-First names rearranged by length:"
For i = 1 To 10
     a = Len(arr1(i))
     e = InStr(1, arr1(i), " ") 'Position of space in the name
     If e > 0 Then 'Skip names with no spaces
        c = Left(arr1(i), e - 1) 'Grab first name
        d = Right(arr1(i), a - e) 'Grab last name
        Call Swap_Strings(c, d)
        Debug.Print c & " " & d
    End If
Next i
   
End Sub

Sub Swap_Strings(Str1 As String, Str2 As String)
Dim a As String

a = Str1
Str1 = Str2
Str2 = a

End Sub

