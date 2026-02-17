// Tests for VBA String Functions
//
// This test file covers VBA string functions including:
// - Len, LenB
// - Left, Right, Mid
// - UCase, LCase
// - Trim, LTrim, RTrim
// - InStr, InStrRev
// - Replace, StrReverse
// - Asc, AscW, Chr, ChrW
// - Space, String
// - StrComp, StrConv
// - Format, FormatCurrency, FormatNumber, FormatPercent

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
// LEN / LENB TESTS
// ============================================================

#[test]
fn test_len_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Len("Hello")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "5");
}

#[test]
fn test_len_empty_string() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Len("")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0");
}

#[test]
fn test_lenb_basic() {
    // LenB returns byte length (UTF-16 = chars * 2)
    let code = r#"
        Sub AutoOpen()
            MsgBox LenB("Hello")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "10");
}

// ============================================================
// LEFT / RIGHT / MID TESTS
// ============================================================

#[test]
fn test_left_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Left("Hello World", 5)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hello");
}

#[test]
fn test_left_zero() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Left("Hello", 0)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "");
}

#[test]
fn test_left_exceeds_length() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Left("Hi", 10)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hi");
}

#[test]
fn test_right_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Right("Hello World", 5)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "World");
}

#[test]
fn test_right_exceeds_length() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Right("Hi", 10)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hi");
}

#[test]
fn test_mid_with_length() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Mid("Hello World", 7, 5)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "World");
}

#[test]
fn test_mid_without_length() {
    // Mid without length returns from start to end
    let code = r#"
        Sub AutoOpen()
            MsgBox Mid("Hello World", 7)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "World");
}

#[test]
fn test_mid_start_at_1() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Mid("Hello", 1, 2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "He");
}

// ============================================================
// UCASE / LCASE TESTS
// ============================================================

#[test]
fn test_ucase_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox UCase("hello world")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "HELLO WORLD");
}

#[test]
fn test_lcase_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox LCase("HELLO WORLD")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "hello world");
}

#[test]
fn test_ucase_mixed() {
    let code = r#"
        Sub AutoOpen()
            MsgBox UCase("HeLLo")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "HELLO");
}

// ============================================================
// TRIM / LTRIM / RTRIM TESTS
// ============================================================

#[test]
fn test_trim_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Trim("  Hello World  ")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hello World");
}

#[test]
fn test_ltrim_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox LTrim("  Hello")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hello");
}

#[test]
fn test_rtrim_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox RTrim("Hello  ")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hello");
}

#[test]
fn test_trim_no_spaces() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Trim("Hello")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hello");
}

// ============================================================
// INSTR / INSTRREV TESTS
// ============================================================

#[test]
fn test_instr_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox InStr("Hello World", "World")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "7");
}

#[test]
fn test_instr_not_found() {
    let code = r#"
        Sub AutoOpen()
            MsgBox InStr("Hello World", "xyz")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0");
}

#[test]
fn test_instr_with_start() {
    let code = r#"
        Sub AutoOpen()
            MsgBox InStr(3, "Hello Hello", "Hello")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "7");
}

#[test]
fn test_instr_first_char() {
    let code = r#"
        Sub AutoOpen()
            MsgBox InStr("Hello", "H")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1");
}

#[test]
fn test_instrrev_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox InStrRev("Hello Hello", "Hello")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "7");
}

// ============================================================
// REPLACE / STRREVERSE TESTS
// ============================================================

#[test]
fn test_replace_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Replace("Hello World", "World", "VBA")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hello VBA");
}

#[test]
fn test_replace_multiple() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Replace("a-b-c-d", "-", "_")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "a_b_c_d");
}

#[test]
fn test_replace_not_found() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Replace("Hello", "xyz", "abc")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hello");
}

#[test]
fn test_strreverse_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox StrReverse("Hello")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "olleH");
}

#[test]
fn test_strreverse_palindrome() {
    let code = r#"
        Sub AutoOpen()
            MsgBox StrReverse("radar")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "radar");
}

// ============================================================
// ASC / ASCW / CHR / CHRW TESTS
// ============================================================

#[test]
fn test_asc_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Asc("A")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "65");
}

#[test]
fn test_asc_lowercase() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Asc("a")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "97");
}

#[test]
fn test_asc_number() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Asc("0")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "48");
}

#[test]
fn test_chr_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Chr(65)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "A");
}

#[test]
fn test_chr_lowercase() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Chr(97)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "a");
}

#[test]
fn test_chr_newline() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Len(Chr(10))
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1");
}

// ============================================================
// SPACE / STRING TESTS
// ============================================================

#[test]
fn test_space_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Len(Space(5))
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "5");
}

#[test]
fn test_space_zero() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Len(Space(0))
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0");
}

#[test]
fn test_string_with_char() {
    let code = r#"
        Sub AutoOpen()
            MsgBox String(5, "x")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "xxxxx");
}

#[test]
fn test_string_with_code() {
    let code = r#"
        Sub AutoOpen()
            MsgBox String(3, 65)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "AAA");
}

// ============================================================
// STRCOMP TESTS
// ============================================================

#[test]
fn test_strcomp_equal() {
    let code = r#"
        Sub AutoOpen()
            MsgBox StrComp("hello", "hello")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0");
}

#[test]
fn test_strcomp_less() {
    let code = r#"
        Sub AutoOpen()
            MsgBox StrComp("abc", "xyz")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "-1");
}

#[test]
fn test_strcomp_greater() {
    let code = r#"
        Sub AutoOpen()
            MsgBox StrComp("xyz", "abc")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1");
}

#[test]
fn test_strcomp_case_sensitive() {
    // Binary compare (default) - case sensitive
    let code = r#"
        Sub AutoOpen()
            MsgBox StrComp("ABC", "abc", 0)
        End Sub
    "#;
    // "ABC" < "abc" in binary comparison
    assert_eq!(run_vba_first(code), "-1");
}

#[test]
fn test_strcomp_case_insensitive() {
    // Text compare - case insensitive
    let code = r#"
        Sub AutoOpen()
            MsgBox StrComp("ABC", "abc", 1)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0");
}

// ============================================================
// STRCONV TESTS
// ============================================================

#[test]
fn test_strconv_uppercase() {
    // vbUpperCase = 1
    let code = r#"
        Sub AutoOpen()
            MsgBox StrConv("hello", 1)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "HELLO");
}

#[test]
fn test_strconv_lowercase() {
    // vbLowerCase = 2
    let code = r#"
        Sub AutoOpen()
            MsgBox StrConv("HELLO", 2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "hello");
}

#[test]
fn test_strconv_propercase() {
    // vbProperCase = 3
    let code = r#"
        Sub AutoOpen()
            MsgBox StrConv("hello world", 3)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hello World");
}

// ============================================================
// FORMAT FUNCTION TESTS
// ============================================================

#[test]
fn test_formatnumber_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox FormatNumber(1234.5678, 2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1234.57");
}

#[test]
fn test_formatnumber_no_decimal() {
    let code = r#"
        Sub AutoOpen()
            MsgBox FormatNumber(1234, 0)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1234");
}

#[test]
fn test_formatpercent_basic() {
    let code = r#"
        Sub AutoOpen()
            MsgBox FormatPercent(0.25, 0)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "25%");
}

#[test]
fn test_formatpercent_with_decimals() {
    let code = r#"
        Sub AutoOpen()
            MsgBox FormatPercent(0.1234, 2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "12.34%");
}

#[test]
fn test_formatcurrency_positive() {
    let code = r#"
        Sub AutoOpen()
            MsgBox FormatCurrency(1234.56)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "$1,234.56");
}

#[test]
fn test_formatcurrency_negative() {
    let code = r#"
        Sub AutoOpen()
            Dim x As Double
            x = -1234.56
            MsgBox FormatCurrency(x)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "-$1,234.56");
}

// ============================================================
// COMBINED STRING OPERATIONS
// ============================================================

#[test]
fn test_combined_trim_ucase() {
    let code = r#"
        Sub AutoOpen()
            MsgBox UCase(Trim("  hello  "))
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "HELLO");
}

#[test]
fn test_combined_left_ucase() {
    let code = r#"
        Sub AutoOpen()
            MsgBox UCase(Left("hello world", 5))
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "HELLO");
}

#[test]
fn test_combined_mid_replace() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Replace(Mid("Hello World", 1, 5), "e", "a")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Hallo");
}

// ============================================================
// EDGE CASES
// ============================================================

#[test]
fn test_empty_string_operations() {
    let code = r#"
        Sub AutoOpen()
            Dim s As String
            s = ""
            MsgBox Len(s)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0");
}

#[test]
fn test_single_char_operations() {
    let code = r#"
        Sub AutoOpen()
            MsgBox Left("X", 1) & Right("X", 1) & Mid("X", 1, 1)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "XXX");
}

// ============================================================
// OPTIONAL PARAMETER TESTS
// ============================================================

// InStr with various calling conventions
#[test]
fn test_instr_two_args() {
    // InStr(string1, string2) - 2 argument form
    let code = r#"
        Sub AutoOpen()
            MsgBox InStr("Hello World", "World")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "7");
}

#[test]
fn test_instr_three_args() {
    // InStr(start, string1, string2) - 3 argument form with start
    let code = r#"
        Sub AutoOpen()
            MsgBox InStr(3, "Hello Hello", "Hello")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "7");
}

#[test]
fn test_instr_case_sensitive() {
    // InStr with compare=0 (default, case-sensitive)
    let code = r#"
        Sub AutoOpen()
            MsgBox InStr(1, "Hello World", "hello", 0)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0");
}

#[test]
fn test_instr_case_insensitive() {
    // InStr with compare=1 (vbTextCompare, case-insensitive)
    let code = r#"
        Sub AutoOpen()
            MsgBox InStr(1, "Hello World", "hello", 1)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1");
}

#[test]
fn test_instr_empty_search_returns_start() {
    // VBA returns start position when search string is empty
    let code = r#"
        Sub AutoOpen()
            MsgBox InStr("Hello", "")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1");
}

// InStrRev with optional parameters
#[test]
fn test_instrrev_with_start() {
    // InStrRev with start position
    let code = r#"
        Sub AutoOpen()
            MsgBox InStrRev("Hello Hello", "Hello", 8)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1");
}

#[test]
fn test_instrrev_case_insensitive() {
    // InStrRev with compare=1 (case-insensitive)
    let code = r#"
        Sub AutoOpen()
            MsgBox InStrRev("Hello World", "WORLD", -1, 1)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "7");
}

// Replace with optional parameters
#[test]
fn test_replace_with_start() {
    // Replace with start position - ignores characters before start
    let code = r#"
        Sub AutoOpen()
            MsgBox Replace("aaaaa", "a", "b", 3)
        End Sub
    "#;
    // Starting at position 3, remaining "aaa" gets replaced
    assert_eq!(run_vba_first(code), "aabbb");
}

#[test]
fn test_replace_with_count() {
    // Replace with count - limits number of replacements
    let code = r#"
        Sub AutoOpen()
            MsgBox Replace("aaaaa", "a", "b", 1, 2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "bbaaa");
}

#[test]
fn test_replace_case_insensitive() {
    // Replace with compare=1 (case-insensitive)
    let code = r#"
        Sub AutoOpen()
            MsgBox Replace("Hello HELLO hello", "hello", "X", 1, -1, 1)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "X X X");
}

#[test]
fn test_replace_all_params() {
    // Replace with all optional parameters
    let code = r#"
        Sub AutoOpen()
            MsgBox Replace("aAaAa", "A", "X", 2, 1, 1)
        End Sub
    "#;
    // Start at pos 2 ("AaAa"), replace 1 occurrence case-insensitively
    assert_eq!(run_vba_first(code), "aXaAa");
}
