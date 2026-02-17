' Test 1: Basic Option Explicit - Should work
Option Explicit

Sub Test1_DeclaredVariable
    Dim x As Integer
    x = 5
    MsgBox x
    MsgBox "Test 1 passed"
End Sub

' Test 2: Undeclared variable - Should fail
Option Explicit

Sub Test2_UndeclaredVariable
    y = 10
    MsgBox y
    MsgBox "This should not appear"
End Sub

' Test 3: Multiple declared variables
Option Explicit

Sub Test3_MultipleDeclared
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    a = 1
    b = 2
    c = 3
    MsgBox a + b + c
    MsgBox "Test 3 passed"
End Sub

' Test 4: Without Option Explicit - Should work
Sub Test4_WithoutExplicit
    z = 100
    MsgBox z
    MsgBox "Test 4 passed - no Option Explicit"
End Sub

' Test 5: Option Explicit with Dim on same line
Option Explicit

Sub Test5_DimMultiple
    Dim x As Integer, y As Integer, z As String
    x = 10
    y = 20
    z = "Hello"
    MsgBox x + y
    MsgBox z
    MsgBox "Test 5 passed"
End Sub

' Test 6: Option Explicit with loops
Option Explicit

Sub Test6_InLoop
    Dim i As Integer
    i = 1
    
    Do While i <= 3
        MsgBox i
        i = i + 1
    Loop
    
    MsgBox "Test 6 passed"
End Sub

' Test 7: Undeclared variable in loop - Should fail
Option Explicit

Sub Test7_UndeclaredInLoop
    Do While counter <= 3
        MsgBox counter
        counter = counter + 1
    Loop
    
    MsgBox "This should not appear"
End Sub

' Test 8: Option Explicit with For loop
Option Explicit

Sub Test8_ForLoop
    Dim i As Integer
    
    For i = 1 To 3
        MsgBox i
    Next i
    
    MsgBox "Test 8 passed"
End Sub

' Test 9: Undeclared For loop counter - Should fail
Option Explicit

Sub Test9_UndeclaredForCounter
    For j = 1 To 3
        MsgBox j
    Next j
    
    MsgBox "This should not appear"
End Sub

' Test 10: Option Explicit with subroutine parameters
Option Explicit

Sub Test10_WithParameters
    Dim x As Integer
    x = 42
    Call Helper10(x)
    MsgBox "Test 10 passed"
End Sub

Sub Helper10(value As Integer)
    MsgBox value
End Sub

' Test 11: Option Explicit with Type
Option Explicit

Type Person
    Name As String
    Age As Integer
End Type

Sub Test11_WithType
    Dim p As Person
    p.Name = "Alice"
    p.Age = 30
    MsgBox p.Name
    MsgBox p.Age
    MsgBox "Test 11 passed"
End Sub

' Test 12: Undeclared Type variable - Should fail
Option Explicit

Type Employee
    ID As Integer
    Name As String
End Type

Sub Test12_UndeclaredType
    emp.ID = 100
    emp.Name = "Bob"
    MsgBox emp.Name
    MsgBox "This should not appear"
End Sub

' Test 13: Option Explicit with nested scopes
Option Explicit

Sub Test13_NestedScopes
    Dim outer As Integer
    outer = 10
    
    Do While outer > 0
        Dim inner As Integer
        inner = outer * 2
        MsgBox inner
        outer = outer - 1
        If outer <= 8 Then
            Exit Do
        End If
    Loop
    
    MsgBox "Test 13 passed"
End Sub

' Test 14: Option Explicit with If statement
Option Explicit

Sub Test14_IfStatement
    Dim x As Integer
    x = 5
    
    If x > 3 Then
        Dim y As Integer
        y = x * 2
        MsgBox y
    Else
        MsgBox "Not executed"
    End If
    
    MsgBox "Test 14 passed"
End Sub

' Test 15: Undeclared in If - Should fail
Option Explicit

Sub Test15_UndeclaredInIf
    Dim x As Integer
    x = 5
    
    If x > 3 Then
        result = x * 2
        MsgBox result
    End If
    
    MsgBox "This should not appear"
End Sub

' Test 16: Option Explicit with error handling
Option Explicit

Sub Test16_WithErrorHandling
    On Error GoTo ErrorHandler
    
    Dim x As Integer
    x = 10
    MsgBox x
    MsgBox "Test 16 passed"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error occurred"
End Sub

' Test 17: String concatenation with Option Explicit
Option Explicit

Sub Test17_StringConcat
    Dim first As String
    Dim last As String
    Dim full As String
    
    first = "John"
    last = "Doe"
    full = first + " " + last
    
    MsgBox full
    MsgBox "Test 17 passed"
End Sub

' Test 18: Undeclared in concatenation - Should fail
Option Explicit

Sub Test18_UndeclaredConcat
    Dim first As String
    first = "Jane"
    result = first + " " + lastName
    MsgBox result
    MsgBox "This should not appear"
End Sub

' Test 19: Boolean with Option Explicit
Option Explicit

Sub Test19_Boolean
    Dim flag As Boolean
    flag = True
    
    Do While flag
        MsgBox "Looping"
        flag = False
    Loop
    
    MsgBox "Test 19 passed"
End Sub

' Test 20: Mixed declared and undeclared - Should fail
Option Explicit

Sub Test20_MixedVars
    Dim x As Integer
    x = 5
    y = 10
    MsgBox x + y
    MsgBox "This should not appear"
End Sub