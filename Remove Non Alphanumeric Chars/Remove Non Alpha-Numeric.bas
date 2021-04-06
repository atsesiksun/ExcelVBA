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
Sub TestRemoveNonAlphaNumeric()

Call RemoveNonAlphaNumeric("12 3nje dfBS<f    678jk;ji#*AAfnj")

End Sub

Sub RemoveNonAlphaNumeric(Str1 As String)
Dim a As Integer
Dim b As String
Dim c As String
Dim x As Integer
Dim Str2 As String

Debug.Print "Question: " & Str1
Debug.Print ("")

Str2 = Mid(Str1, 1, 1) ' To store 1st character and answer

a = Len(Str1) - 1
For x = 1 To a
    b = Mid(Str1, x + 1, 1) 'Access each character in StrVar starting at 2nd one
    c = Mid(Str1, x, 1)  'Access to the character before b
    If Asc(UCase(b)) > 64 And Asc(UCase(b)) < 91 Then 'If character is a letter, add to Str2
        Str2 = Str2 & b
    ElseIf IsNumeric(b) = True And IsNumeric(Right(Str2, 1)) = True Then 'If character is a number and last character in Str2 is a number, add to Str2
        Str2 = Str2 & b
    End If
Next x

Debug.Print "Answer: " & Str2


End Sub



