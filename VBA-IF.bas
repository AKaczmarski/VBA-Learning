Attribute VB_Name = "Module1"
Sub Test()

' If Range("a1") = 1 Then MsgBox "OK"
' If Range("a1") <> 1 Then MsgBox "Error"

If Range("a1") = 1 Then
    MsgBox "OK"
    Range("a2") = "True"

ElseIf Range("a1") = 2 Then
    MsgBox "2"

Else
    MsgBox "Error"
    MsgBox "Change A1 to 1"

End If


End Sub
