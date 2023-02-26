Dim num1
Dim num2
Dim operation
Dim result
num1 = InputBox("Enter the first number:", "Visual Basic Script Calculator")
If num1 = "" Then
    MsgBox "Calculation cancelled.", vbOKOnly, "Visual Basic Script Calculator"
    WScript.Quit
End If
operation = InputBox("Choose an operation:" & vbNewLine & "+ for addition" & vbNewLine & "- for subtraction" & vbNewLine & "x for multiplication" & vbNewLine & "/ for division", "Visual Basic Script Calculator")
If operation = "" Then
    MsgBox "Calculation cancelled.", vbOKOnly, "Visual Basic Script Calculator"
    WScript.Quit
End If
num2 = InputBox("Enter the second number:", "Visual Basic Script Calculator")
If num2 = "" Then
    MsgBox "Calculation cancelled.", vbOKOnly, "Visual Basic Script Calculator"
    WScript.Quit
End If
If IsNumeric(num1) And IsNumeric(num2) Then
    Select Case operation
        Case "+"
            result = CDbl(num1) + CDbl(num2)
        Case "-"
            result = CDbl(num1) - CDbl(num2)
        Case "x"
            result = CDbl(num1) * CDbl(num2)
        Case "/"
            If CDbl(num2) = 0 Then
                MsgBox "Cannot divide by zero.", vbOKOnly, "Visual Basic Script Calculator"
                WScript.Quit
            Else
                result = CDbl(num1) / CDbl(num2)
            End If
        Case Else
            MsgBox "Invalid operation choice.", vbOKOnly, "Visual Basic Script Calculator"
            WScript.Quit
    End Select
    MsgBox "The result is: " & result, vbOKOnly, "Visual Basic Script Calculator"
Else
    MsgBox "Please enter valid numbers.", vbOKOnly, "Visual Basic Script Calculator"
End If