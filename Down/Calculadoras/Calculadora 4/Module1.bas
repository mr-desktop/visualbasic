Attribute VB_Name = "Module1"

Dim Num1, Num2, Result As Single
Dim Operator As String * 1

Sub One() 'Number 1
Form1.Label1.Caption = Form1.Label1.Caption & "1"
End Sub

Sub two()
Form1.Label1.Caption = Form1.Label1.Caption & "2"
End Sub

Sub three()
Form1.Label1.Caption = Form1.Label1.Caption & "3"
End Sub

Sub four()
Form1.Label1.Caption = Form1.Label1.Caption & "4"
End Sub

Sub five()
Form1.Label1.Caption = Form1.Label1.Caption & "5"
End Sub

Sub six()
Form1.Label1.Caption = Form1.Label1.Caption & "6"
End Sub

Sub seven()
Form1.Label1.Caption = Form1.Label1.Caption & "7"
End Sub

Sub eight()
Form1.Label1.Caption = Form1.Label1.Caption & "8"
End Sub

Sub nine()
Form1.Label1.Caption = Form1.Label1.Caption & "9"
End Sub

Sub zero()
Form1.Label1.Caption = Form1.Label1.Caption & "0"
End Sub

Sub clear()
Form1.Label1.Caption = ""
End Sub

Sub Add()
Num1 = Val(Form1.Label1.Caption)
Form1.Label1.Caption = ""
Operator = "+"
End Sub

Sub Equals()

Num2 = Val(Form1.Label1.Caption)

If Operator = "+" Then
    Result = Num1 + Num2
End If

If Operator = "-" Then
    Result = Num1 - Num2
End If

If Operator = "x" Then
    Result = Num1 * Num2
End If

'Stops division of 0 using a logical AND.
If Operator = "÷" And Num2 = 0 Then
    MsgBox "You cannot divide by zero."
    Exit Sub
End If

If Operator = "÷" Then
    Result = Num1 / Num2
End If

Form1.Label1.Caption = Result

'Num1 = 0 'Resets Variables. Uncomment
'Num2 = 0 'these to see what happens.

Call List

End Sub

Sub multiply()
Num1 = Val(Form1.Label1.Caption) 'Converts integer to string.
Form1.Label1.Caption = "" 'Clears label.
Operator = "x"
End Sub

Sub Divide()
Num1 = Val(Form1.Label1.Caption)
Form1.Label1.Caption = ""
Operator = "÷"
End Sub

Sub Subtract()
Num1 = Val(Form1.Label1.Caption)
Form1.Label1.Caption = ""
Operator = "-"
End Sub

Sub List()
'If Result = 0 Then
'    Exit Sub
'End If

Form1.List1.AddItem Result, 0

If Form1.List1.ListCount = 11 Then
    Form1.List1.RemoveItem 10
End If
End Sub

Sub EndIt()
End 'Ends Program.
End Sub

