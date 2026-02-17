// Tests for VBA Err Object Functions
//
// This test file covers VBA Err object functions including:
// - Err.Number - Error number property
// - Err.Description - Error description property
// - Err.Source - Error source property
// - Err.Clear - Clears all error properties
// - Err.Raise - Generates a runtime error

use tree_sitter::Parser;
use vba_parser::language as tree_sitter_vba;
use vba_utils::Context;
use vba_utils::vm::ProgramExecutor;
use vba_utils::ast::build_ast;

/// Helper to run VBA code and capture output
fn run_vba(code: &str) -> Vec<String> {
    let mut parser = Parser::new();
    parser.set_language(tree_sitter_vba()).expect("Failed to set VBA language");
    let tree = parser.parse(code, None).expect("Failed to parse VBA code");
    let root_node = tree.root_node();
    let program = build_ast(root_node, code);
    
    let mut ctx = Context::new();
    let executor = ProgramExecutor::new(program);
    let _ = executor.execute(&mut ctx);
    ctx.output.clone()
}

/// Helper to run VBA code and get first output value
fn run_vba_first(code: &str) -> String {
    let output = run_vba(code);
    output.first().cloned().unwrap_or_default()
}

// ============================================================
// ERR.NUMBER TESTS
// ============================================================

#[test]
fn test_err_number_default() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Err.Number
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0");  // No error, number is 0
}

#[test]
fn test_err_number_after_raise() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(42)
            MsgBox Err.Number
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "42");
}

// ============================================================
// ERR.DESCRIPTION TESTS
// ============================================================

#[test]
fn test_err_description_default() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Err.Description
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "");  // No error, empty description
}

#[test]
fn test_err_description_after_raise() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(13)
            MsgBox Err.Description
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Type mismatch");  // Default description for error 13
}

#[test]
fn test_err_description_custom() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(1001, "MyApp", "Custom error message")
            MsgBox Err.Description
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Custom error message");
}

// ============================================================
// ERR.SOURCE TESTS
// ============================================================

#[test]
fn test_err_source_default() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Err.Source
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "");  // No error, empty source
}

#[test]
fn test_err_source_after_raise() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(13)
            MsgBox Err.Source
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "VBA");  // Default source
}

#[test]
fn test_err_source_custom() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(1001, "MyApp", "Custom error")
            MsgBox Err.Source
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "MyApp");
}

// ============================================================
// ERR.CLEAR TESTS
// ============================================================

#[test]
fn test_err_clear_no_error() {
    let code = r#"
        Sub AutoOpen()
            Err.Clear
            MsgBox Err.Number
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0");
}

#[test]
fn test_err_clear_after_raise() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(13)
            Err.Clear
            MsgBox Err.Number
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0");
}

#[test]
fn test_err_clear_resets_description() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(13, "Test", "Test error")
            Err.Clear
            MsgBox Err.Description
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "");
}

// ============================================================
// ERR.RAISE TESTS
// ============================================================

#[test]
fn test_err_raise_number_only() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(13)
            MsgBox Err.Number
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "13");
}

#[test]
fn test_err_raise_with_source() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(100, "MyModule")
            MsgBox Err.Source
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "MyModule");
}

#[test]
fn test_err_raise_with_all_params() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(1001, "MyApp", "Custom error message")
            MsgBox Err.Number & ":" & Err.Source & ":" & Err.Description
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1001:MyApp:Custom error message");
}

#[test]
fn test_err_raise_default_descriptions() {
    // Test that Err.Raise uses default descriptions for known error numbers
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(11)
            MsgBox Err.Description
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Division by zero");
}

#[test]
fn test_err_raise_overflow_error() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(6)
            MsgBox Err.Description
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Overflow");
}

#[test]
fn test_err_raise_unknown_error() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(99999)
            MsgBox Err.Description
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Application-defined or object-defined error 99999");
}

// ============================================================
// ERR OBJECT WORKFLOW TESTS
// ============================================================

#[test]
fn test_err_workflow_raise_check_clear() {
    // Test that error number is captured, then cleared
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(5)
            MsgBox Err.Number
            Err.Clear
            MsgBox Err.Number
        End Sub
    "#;
    let output = run_vba(code);
    assert_eq!(output.len(), 2);
    assert_eq!(output[0], "5");
    assert_eq!(output[1], "0");
}

#[test]
fn test_err_multiple_raises() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(5)
            Err.Raise(13)
            MsgBox Err.Number
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "13");  // Last error wins
}

#[test]
fn test_err_raise_preserves_custom_message() {
    let code = r#"
        Sub AutoOpen()
            On Error Resume Next
            Err.Raise(1001, "", "First error")
            Err.Raise(1002, "", "Second error")
            MsgBox Err.Description
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Second error");
}
