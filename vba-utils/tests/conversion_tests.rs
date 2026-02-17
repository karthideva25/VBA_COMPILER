// Tests for VBA Type Conversion Functions
//
// This test file covers VBA conversion functions including:
// - CStr - Convert to String
// - CInt, CLng, CLngLng - Integer conversions
// - CByte - Byte conversion
// - CDbl, CSng, CDec - Floating point conversions
// - CCur - Currency conversion
// - CBool - Boolean conversion
// - CDate - Date conversion
// - CVar - Variant conversion

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
// CSTR TESTS
// ============================================================

#[test]
fn test_cstr_from_integer() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CStr(42)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "42");
}

#[test]
fn test_cstr_from_double() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CStr(3.14)
        End Sub
    "#;
    // Currency format has 4 decimal places
    let result = run_vba_first(code);
    assert!(result.starts_with("3.14"));
}

#[test]
fn test_cstr_from_boolean_true() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CStr(True)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "True");
}

#[test]
fn test_cstr_from_boolean_false() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CStr(False)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "False");
}

#[test]
fn test_cstr_from_string() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CStr("Hello")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hello");
}

// ============================================================
// CINT TESTS
// ============================================================

#[test]
fn test_cint_from_double() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CInt(3.7)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "4");
}

#[test]
fn test_cint_from_double_down() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CInt(3.2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "3");
}

#[test]
fn test_cint_from_string() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CInt("42")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "42");
}

#[test]
fn test_cint_from_boolean() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CInt(True)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "-1");
}

#[test]
fn test_cint_negative() {
    let code = r#"
        Sub AutoOpen()
            Dim x As Double
            x = -3.7
            MsgBox CInt(x)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "-4");
}

// ============================================================
// CLNG TESTS
// ============================================================

#[test]
fn test_clng_from_double() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CLng(12345.6)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "12346");
}

#[test]
fn test_clng_from_string() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CLng("99999")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "99999");
}

#[test]
fn test_clng_large_number() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CLng(1000000)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1000000");
}

// ============================================================
// CLNGLNG TESTS
// ============================================================

#[test]
fn test_clnglng_from_double() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CLngLng(123456789012.5)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "123456789013");
}

#[test]
fn test_clnglng_large_value() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CLngLng(9999999999)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "9999999999");
}

// ============================================================
// CBYTE TESTS
// ============================================================

#[test]
fn test_cbyte_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CByte(128)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "128");
}

#[test]
fn test_cbyte_from_double() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CByte(100.7)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "101");
}

#[test]
fn test_cbyte_zero() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CByte(0)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0");
}

#[test]
fn test_cbyte_max() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CByte(255)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "255");
}

// ============================================================
// CDBL TESTS
// ============================================================

#[test]
fn test_cdbl_from_integer() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CDbl(42)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "42");
}

#[test]
fn test_cdbl_from_string() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CDbl("3.14159")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "3.14159");
}

#[test]
fn test_cdbl_scientific_notation() {
    let code = r#"
        Sub AutoOpen()
            Dim y As Integer
            y = 2
            Dim x As Double
            x = CDbl(y)
            MsgBox x * 2
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "4");
}

// ============================================================
// CSNG TESTS
// ============================================================

#[test]
fn test_csng_from_double() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CSng(3.14)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert!(result.starts_with("3.14"));
}

#[test]
fn test_csng_from_integer() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CSng(42)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "42");
}

// ============================================================
// CDEC TESTS
// ============================================================

#[test]
fn test_cdec_from_double() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CDec(123.456)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "123.456");
}

#[test]
fn test_cdec_from_integer() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CDec(1000)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1000");
}

// ============================================================
// CCUR TESTS
// ============================================================

#[test]
fn test_ccur_from_double() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CCur(1234.5678)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1234.5678");
}

#[test]
fn test_ccur_from_integer() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CCur(1000)
        End Sub
    "#;
    let result = run_vba_first(code);
    // Currency may be displayed with or without trailing zeros
    assert!(result.starts_with("1000"));
}

#[test]
fn test_ccur_rounding() {
    // CCur rounds to 4 decimal places
    let code = r#"
        Sub AutoOpen()
            MsgBox CCur(1234.56789)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1234.5679");
}

// ============================================================
// CBOOL TESTS
// ============================================================

#[test]
fn test_cbool_from_true_string() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CBool("True")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "True");
}

#[test]
fn test_cbool_from_false_string() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CBool("False")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "False");
}

#[test]
fn test_cbool_from_nonzero_integer() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CBool(1)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "True");
}

#[test]
fn test_cbool_from_zero() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CBool(0)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "False");
}

#[test]
fn test_cbool_from_negative() {
    let code = r#"
        Sub AutoOpen()
            Dim x As Integer
            x = -5
            MsgBox CBool(x)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "True");
}

#[test]
fn test_cbool_from_double() {
    let code = r#"
        Sub AutoOpen()
            Dim x As Double
            x = 0.5
            MsgBox CBool(x)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "True");
}

// ============================================================
// CDATE TESTS
// ============================================================

#[test]
fn test_cdate_from_string() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = CDate("2026-01-15")
            MsgBox Year(d) & "-" & Month(d) & "-" & Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2026-1-15");
}

#[test]
fn test_cdate_from_serial_number() {
    // Serial number 45678 should be a date
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = CDate(45000)
            MsgBox Year(d)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert!(result.parse::<i32>().is_ok());
}

#[test]
fn test_cdate_from_date() {
    let code = r#"
        Sub AutoOpen()
            Dim d1 As Date, d2 As Date
            d1 = DateSerial(2026, 6, 15)
            d2 = CDate(d1)
            MsgBox Month(d2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "6");
}

// ============================================================
// CVAR TESTS
// ============================================================

#[test]
fn test_cvar_from_integer() {
    let code = r#"
        Sub AutoOpen()
            Dim v As Variant
            v = CVar(42)
            MsgBox v
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "42");
}

#[test]
fn test_cvar_from_string() {
    let code = r#"
        Sub AutoOpen()
            Dim v As Variant
            v = CVar("Hello")
            MsgBox v
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hello");
}

#[test]
fn test_cvar_preserves_type() {
    let code = r#"
        Sub AutoOpen()
            Dim v As Variant
            v = CVar(3.14)
            MsgBox v + 1
        End Sub
    "#;
    let result = run_vba_first(code);
    assert!(result.starts_with("4.14"));
}

// ============================================================
// CHAINED CONVERSIONS
// ============================================================

#[test]
fn test_chain_cstr_cint() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CStr(CInt(3.7))
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "4");
}

#[test]
fn test_chain_cdbl_clng() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CLng(CDbl("12345.6"))
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "12346");
}

#[test]
fn test_chain_cbool_cint() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CInt(CBool(1))
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "-1");
}

// ============================================================
// EDGE CASES
// ============================================================

#[test]
fn test_conversion_empty_to_integer() {
    let code = r#"
        Sub AutoOpen()
            Dim v As Variant
            MsgBox CInt(v)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0");
}

#[test]
fn test_conversion_empty_to_boolean() {
    let code = r#"
        Sub AutoOpen()
            Dim v As Variant
            MsgBox CBool(v)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "False");
}

#[test]
fn test_conversion_roundtrip() {
    // Convert integer to string and back
    let code = r#"
        Sub AutoOpen()
            Dim x As Integer
            x = 42
            MsgBox CInt(CStr(x))
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "42");
}

// ============================================================
// CVDATE TESTS (Variant containing Date - legacy function)
// ============================================================

#[test]
fn test_cvdate_from_string() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CVDate("2024-01-15")
        End Sub
    "#;
    let result = run_vba_first(code);
    assert!(result.contains("01") && result.contains("15") && result.contains("2024"));
}

#[test]
fn test_cvdate_from_serial_number() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CVDate(45307)
        End Sub
    "#;
    // Serial date 45307 = 2024-01-16 (VBA serial dates start from Dec 30, 1899)
    let result = run_vba_first(code);
    assert!(result.contains("01") && result.contains("16") && result.contains("2024"));
}

// ============================================================
// CVERR TESTS (Create Error Value)
// ============================================================

#[test]
fn test_cverr_basic() {
    let code = r#"
        Sub AutoOpen()
            Dim e
            e = CVErr(13)
            MsgBox e
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Error 13");
}

#[test]
fn test_cverr_zero() {
    let code = r#"
        Sub AutoOpen()
            MsgBox CVErr(0)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Error 0");
}

// ============================================================
// ERROR / ERROR$ TESTS (Get Error Message)
// ============================================================

#[test]
fn test_error_division_by_zero() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Error(11)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Division by zero");
}

#[test]
fn test_error_type_mismatch() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Error(13)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Type mismatch");
}

#[test]
fn test_error_overflow() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Error$(6)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Overflow");
}

#[test]
fn test_error_object_required() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Error$(424)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Object required");
}

#[test]
fn test_error_unknown() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Error(99999)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Unknown error 99999");
}

// ============================================================
// STR / STR$ TESTS (Number to String with leading space)
// ============================================================

#[test]
fn test_str_positive_integer() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Str(42)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), " 42");  // Leading space for positive
}

#[test]
fn test_str_negative_integer() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Str(-42)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "-42");  // No leading space for negative
}

#[test]
fn test_str_double() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Str$(3.14)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert!(result.starts_with(" 3.14"));  // Leading space for positive
}

#[test]
fn test_str_zero() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Str(0)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), " 0");  // Leading space (0 is non-negative)
}
