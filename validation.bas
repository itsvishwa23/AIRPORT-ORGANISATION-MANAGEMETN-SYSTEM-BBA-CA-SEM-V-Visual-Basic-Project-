Attribute VB_Name = "validation"
Dim j As Integer




Public Function character(KeyAscii As Integer) As Integer
Dim var  As Boolean
var = chr(KeyAscii) Like "[a-z A-Z]"
If var = False And KeyAscii <> 8 Then
MsgBox "PLEASE ENTER ALPHABETS ONLY"
KeyAscii = 0
End If
character = KeyAscii

End Function

Public Function number(KeyAscii As Integer) As Integer
Dim stat As String
stat = "0123456789"
If InStr(stat, chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
MsgBox "Please enter only number's"
KeyAscii = 0
End If
number = KeyAscii
End Function

Public Function address(KeyAscii As Integer) As Integer
Dim Add  As Boolean
Add = chr(KeyAscii) Like "[""a-z A-Z 0-9,.\ /]"
If Add = False And KeyAscii <> 8 Then
KeyAscii = 0
MsgBox "PLEASE ENTER ALPHABETS & NUMBERS ONLY"
End If
address = KeyAscii
End Function



Public Function numchar(KeyAscii As Integer) As Integer
Dim Add  As Boolean
Add = chr(KeyAscii) Like "[a-z A-Z 0-9]"
If Add = False And KeyAscii <> 8 Then
KeyAscii = 0
End If
numchar = KeyAscii
End Function

Public Function noletter(KeyAscii As Integer) As Integer
KeyAscii = 0
noletter = KeyAscii
End Function
Public Function numcharsymbol(KeyAscii As Integer) As Integer
Dim Add  As Boolean
Add = chr(KeyAscii) Like "[a-z A-Z 0-9 @ . _]"
If Add = False And KeyAscii <> 8 Then
KeyAscii = 0
End If
numcharsymbol = KeyAscii
End Function

Public Function mobnumber(KeyAscii As Integer) As Integer
Dim stat As String
stat = "0123456789"
i = Len(chr(KeyAscii))
If InStr(stat, chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
KeyAscii = 0
End If
i = Len(chr(KeyAscii))

'If i > 10 Then
'MsgBox "Mobile number should contain 10 digits!", vbInformation, "Error"
'Exit Function
'End If


mobnumber = KeyAscii


End Function

Public Function valemail(mail As String) As Boolean
If Left(mail, 1) = "@" Then
GoTo Notvalidet
ElseIf Right(mail, 1) = "@" Then
GoTo Notvalidet
ElseIf InStr(1, mail, "@") = False Then
MsgBox "the @ is missing!"
ElseIf InStr(1, mail, ".") = False Then
GoTo Notvalidet
ElseIf Right(mail, 1) = "." Then
GoTo Notvalidet
ElseIf Left(mail, 1) = "." Then
GoTo Notvalidet
Notvalidet:
MsgBox "This is Not a valid Email address!"
'valemail = False
End If
End Function


