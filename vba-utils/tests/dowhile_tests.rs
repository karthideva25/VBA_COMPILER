#[cfg(test)]
mod do_while_tests {
    use crate::parser::parse_vba;
    use crate::interpreter::execute_program;
    use crate::context::Context;

    #[test]
    fn test_do_while_basic() {
        let code = r#"
Sub TestDoWhile
    Dim i As Integer
    i = 1
    
    Do While i <= 3
        MsgBox i
        i = i + 1
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Done");
    }

    #[test]
    fn test_do_while_zero_iterations() {
        let code = r#"
Sub TestDoWhile
    Dim i As Integer
    i = 10
    
    Do While i <= 3
        MsgBox i
        i = i + 1
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        // Should only output "Done" because condition is false from start
        assert_eq!(ctx.output.len(), 1);
        assert_eq!(ctx.output[0], "Done");
    }

    #[test]
    fn test_do_until_basic() {
        let code = r#"
Sub TestDoUntil
    Dim i As Integer
    i = 1
    
    Do Until i > 3
        MsgBox i
        i = i + 1
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Done");
    }

    #[test]
    fn test_do_loop_while_post_test() {
        let code = r#"
Sub TestDoLoopWhile
    Dim i As Integer
    i = 10
    
    Do
        MsgBox i
        i = i + 1
    Loop While i <= 3
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        // Should execute once even though condition is false
        assert_eq!(ctx.output.len(), 2);
        assert_eq!(ctx.output[0], "10");
        assert_eq!(ctx.output[1], "Done");
    }

    #[test]
    fn test_do_loop_until_post_test() {
        let code = r#"
Sub TestDoLoopUntil
    Dim i As Integer
    i = 1
    
    Do
        MsgBox i
        i = i + 1
    Loop Until i > 3
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Done");
    }

    #[test]
    fn test_do_loop_infinite_with_exit() {
        let code = r#"
Sub TestInfiniteLoop
    Dim i As Integer
    i = 1
    
    Do
        MsgBox i
        i = i + 1
        If i > 3 Then
            Exit Do
        End If
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Done");
    }

    #[test]
    fn test_exit_do_basic() {
        let code = r#"
Sub TestExitDo
    Dim i As Integer
    i = 1
    
    Do While i <= 10
        MsgBox i
        If i = 3 Then
            Exit Do
        End If
        i = i + 1
    Loop
    
    MsgBox "Exited early"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Exited early");
    }

    #[test]
    fn test_nested_do_loops() {
        let code = r#"
Sub TestNestedLoops
    Dim i As Integer
    Dim j As Integer
    i = 1
    
    Do While i <= 2
        MsgBox i
        j = 1
        Do While j <= 2
            MsgBox j
            j = j + 1
        Loop
        i = i + 1
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 7);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "1");
        assert_eq!(ctx.output[2], "2");
        assert_eq!(ctx.output[3], "2");
        assert_eq!(ctx.output[4], "1");
        assert_eq!(ctx.output[5], "2");
        assert_eq!(ctx.output[6], "Done");
    }

    #[test]
    fn test_do_while_with_goto() {
        let code = r#"
Sub TestDoWhileGoTo
    Dim i As Integer
    i = 1
    
    Do While i <= 5
        MsgBox i
        If i = 3 Then
            GoTo ExitPoint
        End If
        i = i + 1
    Loop
    
    MsgBox "Should not reach here"
    
ExitPoint:
    MsgBox "Jumped out"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Jumped out");
    }

    #[test]
    fn test_do_while_on_error_resume_next() {
        let code = r#"
Sub TestDoWhileErrorHandling
    On Error Resume Next
    
    Dim i As Integer
    i = 1
    
    Do While i <= 3
        MsgBox i
        i = i + 1
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Done");
    }

    #[test]
    fn test_do_while_on_error_goto() {
        let code = r#"
Sub TestDoWhileErrorGoTo
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    i = 1
    
    Do While i <= 5
        MsgBox i
        If i = 3 Then
            GoTo ErrorHandler
        End If
        i = i + 1
    Loop
    
    MsgBox "Should not reach"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error handled"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Error handled");
    }

    #[test]
    fn test_do_while_boolean_condition() {
        let code = r#"
Sub TestBooleanCondition
    Dim keepGoing As Boolean
    Dim i As Integer
    keepGoing = True
    i = 1
    
    Do While keepGoing
        MsgBox i
        i = i + 1
        If i > 3 Then
            keepGoing = False
        End If
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Done");
    }

    #[test]
    fn test_do_until_boolean_condition() {
        let code = r#"
Sub TestBooleanUntil
    Dim stopNow As Boolean
    Dim i As Integer
    stopNow = False
    i = 1
    
    Do Until stopNow
        MsgBox i
        i = i + 1
        If i > 3 Then
            stopNow = True
        End If
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Done");
    }

    #[test]
    fn test_do_while_string_comparison() {
        let code = r#"
Sub TestStringCondition
    Dim text As String
    Dim i As Integer
    text = "go"
    i = 1
    
    Do While text = "go"
        MsgBox i
        i = i + 1
        If i > 3 Then
            text = "stop"
        End If
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Done");
    }

    #[test]
    fn test_do_while_complex_expression() {
        let code = r#"
Sub TestComplexCondition
    Dim i As Integer
    Dim j As Integer
    i = 1
    j = 10
    
    Do While i + j <= 15
        MsgBox i
        i = i + 1
        j = j - 1
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Done");
    }

    #[test]
    fn test_exit_do_from_nested_loop() {
        let code = r#"
Sub TestExitNestedDo
    Dim i As Integer
    Dim j As Integer
    i = 1
    
    Do While i <= 3
        MsgBox i
        j = 1
        Do While j <= 3
            MsgBox j
            If j = 2 Then
                Exit Do
            End If
            j = j + 1
        Loop
        i = i + 1
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        // Each outer iteration: i, 1, 2, then exits inner
        assert_eq!(ctx.output.len(), 10);
        assert_eq!(ctx.output[0], "1"); // outer i
        assert_eq!(ctx.output[1], "1"); // inner j
        assert_eq!(ctx.output[2], "2"); // inner j, then exit
        assert_eq!(ctx.output[3], "2"); // outer i
        assert_eq!(ctx.output[4], "1"); // inner j
        assert_eq!(ctx.output[5], "2"); // inner j, then exit
        assert_eq!(ctx.output[6], "3"); // outer i
        assert_eq!(ctx.output[7], "1"); // inner j
        assert_eq!(ctx.output[8], "2"); // inner j, then exit
        assert_eq!(ctx.output[9], "Done");
    }

    #[test]
    fn test_do_while_counter_decrement() {
        let code = r#"
Sub TestCountDown
    Dim i As Integer
    i = 3
    
    Do While i > 0
        MsgBox i
        i = i - 1
    Loop
    
    MsgBox "Blastoff"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "3");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "1");
        assert_eq!(ctx.output[3], "Blastoff");
    }

    #[test]
    fn test_do_while_with_label_inside() {
        let code = r#"
Sub TestLabelInLoop
    Dim i As Integer
    i = 1
    
    Do While i <= 3
LoopStart:
        MsgBox i
        i = i + 1
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Done");
    }

    #[test]
    fn test_do_until_with_not_operator() {
        let code = r#"
Sub TestUntilNot
    Dim done As Boolean
    Dim i As Integer
    done = False
    i = 1
    
    Do Until done
        MsgBox i
        i = i + 1
        If i > 3 Then
            done = True
        End If
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 4);
        assert_eq!(ctx.output[0], "1");
        assert_eq!(ctx.output[1], "2");
        assert_eq!(ctx.output[2], "3");
        assert_eq!(ctx.output[3], "Done");
    }

    #[test]
    fn test_do_loop_empty_body() {
        let code = r#"
Sub TestEmptyBody
    Dim i As Integer
    i = 5
    
    Do While i < 3
    Loop
    
    MsgBox "Done"
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        assert_eq!(ctx.output.len(), 1);
        assert_eq!(ctx.output[0], "Done");
    }

    #[test]
    fn test_do_loop_while_empty_body_executes_once() {
        let code = r#"
Sub TestEmptyBodyPostTest
    Dim i As Integer
    i = 5
    
    Do
        i = i + 1
    Loop While i < 3
    
    MsgBox i
End Sub
"#;
        let mut ctx = Context::default();
        let program = parse_vba(code).expect("Parse failed");
        execute_program(&program, &mut ctx);
        
        // Should execute once, incrementing i from 5 to 6
        assert_eq!(ctx.output.len(), 1);
        assert_eq!(ctx.output[0], "6");
    }
}