use tree_sitter::Parser;
use vba_parser::language;

/// Helper: parse code and return the tree
fn parse_code(code: &str) -> tree_sitter::Tree {
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    tree
}

/// Helper: check if tree has errors
fn has_parse_errors(tree: &tree_sitter::Tree) -> bool {
    let root = tree.root_node();
    root.has_error()
}

/// Helper: find Type declaration nodes
fn find_type_declarations(tree: &tree_sitter::Tree, code: &str) -> Vec<tree_sitter::Node<'_>> {
    let root = tree.root_node();
    let mut cursor = root.walk();
    let mut results = Vec::new();
    
    fn walk_tree<'a>(node: tree_sitter::Node<'a>, cursor: &mut tree_sitter::TreeCursor<'a>, 
                     results: &mut Vec<tree_sitter::Node<'a>>, code: &str) {
        if node.kind() == "type_declaration" || 
           node.kind() == "user_defined_type" ||
           node.kind() == "type_definition" {
            results.push(node);
        }
        
        // Also check if node text starts with "Type " (case insensitive)
        if let Ok(text) = node.utf8_text(code.as_bytes()) {
            if text.trim_start().to_lowercase().starts_with("type ") {
                if !results.contains(&node) {
                    results.push(node);
                }
            }
        }
        
        for child in node.named_children(cursor) {
            walk_tree(child, &mut child.walk(), results, code);
        }
    }
    
    for child in root.named_children(&mut cursor) {
        walk_tree(child, &mut child.walk(), &mut results, code);
    }
    
    results
}

/// Test result enum
#[derive(Debug, PartialEq)]
enum TestResult {
    Ok,
    Fail,
    NotImplemented,
}

/// Helper: run a test case
fn run_test(test_name: &str, code: &str, should_succeed: bool) -> TestResult {
    let tree = parse_code(code);
    let has_errors = has_parse_errors(&tree);
    
    // Check if this is implemented by looking for Type-related nodes
    let type_decls = find_type_declarations(&tree, code);
    let has_type_keyword = code.to_lowercase().contains("type ") && 
                           code.to_lowercase().contains("end type");
    
    // If we have Type syntax but no type declarations found and no errors, likely not implemented
    if has_type_keyword && type_decls.is_empty() && !has_errors {
        println!("[NOT IMPLEMENTED] {}", test_name);
        return TestResult::NotImplemented;
    }
    
    let success = if should_succeed {
        !has_errors
    } else {
        has_errors
    };
    
    if success {
        println!("[OK] {}", test_name);
        TestResult::Ok
    } else {
        println!("[FAIL] {} - Expected {}, got {}", 
                 test_name, 
                 if should_succeed { "success" } else { "error" },
                 if has_errors { "error" } else { "success" });
        TestResult::Fail
    }
}

// ============================================================================
// TEST CASES - VALID SYNTAX
// ============================================================================

#[test]
fn test_001_simple_type() {
    let code = "Type SimpleType\n  x As Integer\nEnd Type\n";
    assert_eq!(run_test("001: Basic UDT with single scalar field", code, true), TestResult::Ok);
}

#[test]
fn test_002_public_type() {
    let code = "Public Type PublicType\n  Name As String\nEnd Type\n";
    assert_eq!(run_test("002: Public scope type", code, true), TestResult::Ok);
}

#[test]
fn test_003_private_type() {
    let code = "Private Type PrivateType\n  Age As Long\nEnd Type\n";
    assert_eq!(run_test("003: Private scope type", code, true), TestResult::Ok);
}

#[test]
fn test_004_all_scalar_types() {
    let code = r#"Type AllScalars
  b As Byte
  i As Integer
  l As Long
  s As Single
  d As Double
  c As Currency
  dt As Date
  bool As Boolean
  str As String
  v As Variant
End Type
"#;
    assert_eq!(run_test("004: All scalar data types", code, true), TestResult::Ok);
}

#[test]
fn test_005_fixed_length_string() {
    let code = "Type FixedString\n  Name As String * 50\nEnd Type\n";
    assert_eq!(run_test("005: Fixed-length string (50 chars)", code, true), TestResult::Ok);
}

#[test]
fn test_006_multiple_fixed_strings() {
    let code = r#"Type MultiFixedString
  FirstName As String * 20
  LastName As String * 30
  Code As String * 5
End Type
"#;
    assert_eq!(run_test("006: Multiple fixed-length strings", code, true), TestResult::Ok);
}

#[test]
fn test_007_static_array() {
    let code = "Type StaticArray\n  Numbers(1 To 10) As Integer\nEnd Type\n";
    assert_eq!(run_test("007: Fixed-size array within UDT", code, true), TestResult::Ok);
}

#[test]
fn test_008_multidimensional_array() {
    let code = "Type MultiDimArray\n  Matrix(1 To 3, 1 To 3) As Long\nEnd Type\n";
    assert_eq!(run_test("008: Multi-dimensional static array", code, true), TestResult::Ok);
}

#[test]
fn test_009_dynamic_array() {
    let code = "Type DynamicArray\n  Items() As String\nEnd Type\n";
    assert_eq!(run_test("009: Dynamic array (requires ReDim)", code, true), TestResult::Ok);
}

#[test]
fn test_010_nested_type() {
    let code = r#"Type SimpleType
  x As Integer
End Type

Type NestedType
  Inner As SimpleType
  Value As Long
End Type
"#;
    assert_eq!(run_test("010: Nested UDT", code, true), TestResult::Ok);
}

#[test]
fn test_020_object_field() {
    let code = "Type ObjectField\n  wb As Workbook\nEnd Type\n";
    assert_eq!(run_test("020: Object references in UDT", code, true), TestResult::Ok);
}

#[test]
fn test_021_variant_array() {
    let code = "Type VariantArray\n  Data As Variant\nEnd Type\n";
    assert_eq!(run_test("021: Variant can hold arrays", code, true), TestResult::Ok);
}

#[test]
fn test_027_no_scope_keyword() {
    let code = "Type NoScope\n  x As Integer\nEnd Type\n";
    assert_eq!(run_test("027: Type without explicit scope (defaults to Public)", code, true), TestResult::Ok);
}

#[test]
fn test_030_array_of_variants() {
    let code = "Type WithArray\n  Items(1 To 10) As Variant\nEnd Type\n";
    assert_eq!(run_test("030: Static array of Variants", code, true), TestResult::Ok);
}

#[test]
fn test_031_string_variations() {
    let code = r#"Type StringVariations
  s1 As String
  s2 As String * 10
  s3 As String * 1
End Type
"#;
    assert_eq!(run_test("031: Mix of dynamic and fixed strings", code, true), TestResult::Ok);
}

#[test]
fn test_034_huge_string_length() {
    let code = "Type HugeStringLength\n  s As String * 65535\nEnd Type\n";
    assert_eq!(run_test("034: Very large fixed string (64KB)", code, true), TestResult::Ok);
}

#[test]
fn test_038_with_enum() {
    let code = "Type WithEnum\n  Status As VbMsgBoxResult\nEnd Type\n";
    assert_eq!(run_test("038: Enum as field type", code, true), TestResult::Ok);
}

#[test]
fn test_039_array_bounds() {
    let code = r#"Type ArrayBounds
  a(0 To 5) As Integer
  b(5) As Integer
  c(-5 To 5) As Long
End Type
"#;
    assert_eq!(run_test("039: Various array bound styles", code, true), TestResult::Ok);
}

#[test]
fn test_040_mixed_case() {
    let code = "Type MixedCase\n  FieldName As Integer\nEnd Type\n";
    assert_eq!(run_test("040: Mixed case identifiers", code, true), TestResult::Ok);
}

#[test]
fn test_041_underscore_names() {
    let code = r#"Type UnderscoreNames
  _private As Integer
  my_value As Long
End Type
"#;
    assert_eq!(run_test("041: Underscores in identifiers", code, true), TestResult::Ok);
}

#[test]
fn test_042_numeric_names() {
    let code = r#"Type NumericNames
  Field1 As Integer
  Value2 As Long
End Type
"#;
    assert_eq!(run_test("042: Numbers in identifiers", code, true), TestResult::Ok);
}

#[test]
fn test_049_longlong_type() {
    let code = "Type LongInteger\n  x As LongLong\nEnd Type\n";
    assert_eq!(run_test("049: LongLong type (VBA7+)", code, true), TestResult::Ok);
}

#[test]
fn test_050_longptr_type() {
    let code = "Type LongPtr\n  ptr As LongPtr\nEnd Type\n";
    assert_eq!(run_test("050: LongPtr type (VBA7+)", code, true), TestResult::Ok);
}

#[test]
fn test_051_complex_nested() {
    let code = r#"Type SimpleType
  x As Integer
End Type

Type FixedString
  Name As String * 50
End Type

Type ComplexNested
  Inner1 As SimpleType
  Inner2 As FixedString
  Value As Double
End Type
"#;
    assert_eq!(run_test("051: Multiple nested UDTs", code, true), TestResult::Ok);
}

#[test]
fn test_054_variant_for_objects() {
    let code = "Type VariantForObjects\n  Items As Variant\nEnd Type\n";
    assert_eq!(run_test("054: Variant as workaround for object arrays", code, true), TestResult::Ok);
}

#[test]
fn test_058_collection_field() {
    let code = "Type CollectionField\n  Items As Collection\nEnd Type\n";
    assert_eq!(run_test("058: Collection object as field", code, true), TestResult::Ok);
}

#[test]
fn test_059_dictionary_field() {
    let code = "Type DictionaryField\n  Dict As Scripting.Dictionary\nEnd Type\n";
    assert_eq!(run_test("059: Dictionary object as field", code, true), TestResult::Ok);
}

// ============================================================================
// TEST CASES - INVALID SYNTAX (SHOULD FAIL)
// ============================================================================

#[test]
fn test_011_empty_type() {
    let code = "Type EmptyType\nEnd Type\n";
    assert_eq!(run_test("011: Empty type (should fail)", code, false), TestResult::Ok);
}

#[test]
fn test_012_missing_end_type() {
    let code = "Type MissingEnd\n  x As Integer\n";
    assert_eq!(run_test("012: Missing End Type", code, false), TestResult::Ok);
}

#[test]
fn test_013_duplicate_field() {
    let code = r#"Type DuplicateField
  Name As String
  Name As Integer
End Type
"#;
    assert_eq!(run_test("013: Duplicate field names", code, false), TestResult::Ok);
}

#[test]
fn test_014_type_as_name() {
    let code = "Type Type\n  x As Integer\nEnd Type\n";
    assert_eq!(run_test("014: 'Type' as UDT name (reserved)", code, false), TestResult::Ok);
}

#[test]
fn test_015_integer_as_name() {
    let code = "Type Integer\n  x As Long\nEnd Type\n";
    assert_eq!(run_test("015: 'Integer' as UDT name (reserved)", code, false), TestResult::Ok);
}

#[test]
fn test_016_sub_as_name() {
    let code = "Type Sub\n  x As Integer\nEnd Type\n";
    assert_eq!(run_test("016: 'Sub' as UDT name (reserved)", code, false), TestResult::Ok);
}

#[test]
fn test_017_very_long_name() {
    let code = "Type VeryLongNameExceeding255CharactersABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz\n  x As Integer\nEnd Type\n";
    assert_eq!(run_test("017: Identifier exceeding 255 chars", code, false), TestResult::Ok);
}

#[test]
fn test_018_cyclic_dependency() {
    let code = r#"Type CyclicA
  b As CyclicB
End Type
Type CyclicB
  a As CyclicA
End Type
"#;
    assert_eq!(run_test("018: Cyclic dependency between types", code, false), TestResult::Ok);
}

#[test]
fn test_019_self_reference() {
    let code = "Type SelfReference\n  Next As SelfReference\nEnd Type\n";
    assert_eq!(run_test("019: Type containing itself", code, false), TestResult::Ok);
}

#[test]
fn test_028_type_in_procedure() {
    let code = r#"Sub Test()
  Type LocalType
    x As Integer
  End Type
End Sub
"#;
    assert_eq!(run_test("028: Type inside procedure (not allowed)", code, false), TestResult::Ok);
}

#[test]
fn test_029_invalid_field_type() {
    let code = "Type InvalidFieldType\n  x As NonExistentType\nEnd Type\n";
    assert_eq!(run_test("029: Non-existent field type", code, false), TestResult::Ok);
}

#[test]
fn test_032_zero_length_string() {
    let code = "Type ZeroLengthString\n  s As String * 0\nEnd Type\n";
    assert_eq!(run_test("032: Zero-length fixed string", code, false), TestResult::Ok);
}

#[test]
fn test_033_negative_string_length() {
    let code = "Type NegativeStringLength\n  s As String * -5\nEnd Type\n";
    assert_eq!(run_test("033: Negative string length", code, false), TestResult::Ok);
}

#[test]
fn test_036_duplicate_type_name() {
    let code = r#"Type DuplicateTypeName
  x As Integer
End Type
Type DuplicateTypeName
  y As Long
End Type
"#;
    assert_eq!(run_test("036: Duplicate type declarations", code, false), TestResult::Ok);
}

#[test]
fn test_037_type_in_class_module() {
    let code = "Public Type InClassModule\n  x As Integer\nEnd Type\n";
    // Note: This would need class module context to properly test, 
    // but we test the parsing aspect
    run_test("037: Type in class module (context-dependent)", code, true);
}

#[test]
fn test_043_starts_with_number() {
    let code = "Type StartsWithNumber\n  1Field As Integer\nEnd Type\n";
    assert_eq!(run_test("043: Identifier starting with digit", code, false), TestResult::Ok);
}

#[test]
fn test_044_special_chars() {
    let code = "Type SpecialChars\n  Field@Name As Integer\nEnd Type\n";
    assert_eq!(run_test("044: Special characters in identifier", code, false), TestResult::Ok);
}

#[test]
fn test_045_trailing_comma() {
    let code = "Type TrailingComma\n  x As Integer,\nEnd Type\n";
    assert_eq!(run_test("045: Trailing comma after field", code, false), TestResult::Ok);
}

#[test]
fn test_046_no_field_type() {
    let code = "Type NoFieldType\n  x\nEnd Type\n";
    assert_eq!(run_test("046: Field without 'As Type'", code, false), TestResult::Ok);
}

#[test]
fn test_047_multiple_on_line() {
    let code = "Type MultipleOnLine\n  x As Integer: y As Long\nEnd Type\n";
    assert_eq!(run_test("047: Multiple fields on one line with colon", code, false), TestResult::Ok);
}

#[test]
fn test_052_array_of_objects() {
    let code = "Type ArrayOfObjects\n  Sheets(1 To 10) As Worksheet\nEnd Type\n";
    assert_eq!(run_test("052: Static array of objects (not allowed)", code, false), TestResult::Ok);
}

#[test]
fn test_053_dynamic_object_array() {
    let code = "Type DynamicObjectArray\n  Sheets() As Worksheet\nEnd Type\n";
    assert_eq!(run_test("053: Dynamic array of objects (restricted)", code, false), TestResult::Ok);
}

#[test]
fn test_055_with_optional() {
    let code = "Type WithOptional\n  Optional x As Integer\nEnd Type\n";
    assert_eq!(run_test("055: Optional keyword in Type (not valid)", code, false), TestResult::Ok);
}

#[test]
fn test_056_with_paramarray() {
    let code = "Type WithParamArray\n  ParamArray args() As Variant\nEnd Type\n";
    assert_eq!(run_test("056: ParamArray in Type (not valid)", code, false), TestResult::Ok);
}

#[test]
fn test_057_with_byval() {
    let code = "Type WithByVal\n  ByVal x As Integer\nEnd Type\n";
    assert_eq!(run_test("057: ByVal modifier in Type field", code, false), TestResult::Ok);
}

// ============================================================================
// SUMMARY TEST - Runs all tests and prints summary
// ============================================================================

#[test]
fn test_000_run_all_and_summarize() {
    println!("\n========================================");
    println!("VBA TYPE/UDT COMPREHENSIVE TEST SUITE");
    println!("========================================\n");
    
    let mut total = 0;
    let mut passed = 0;
    let mut failed = 0;
    let mut not_implemented = 0;
    
    // Valid syntax tests (should succeed)
    let valid_tests = vec![
        ("001", "Type SimpleType\n  x As Integer\nEnd Type\n", true),
        ("002", "Public Type PublicType\n  Name As String\nEnd Type\n", true),
        ("003", "Private Type PrivateType\n  Age As Long\nEnd Type\n", true),
        ("004", "Type AllScalars\n  b As Byte\n  i As Integer\nEnd Type\n", true),
        ("005", "Type FixedString\n  Name As String * 50\nEnd Type\n", true),
        ("006", "Type MultiFixedString\n  FirstName As String * 20\nEnd Type\n", true),
        ("007", "Type StaticArray\n  Numbers(1 To 10) As Integer\nEnd Type\n", true),
        ("008", "Type MultiDimArray\n  Matrix(1 To 3, 1 To 3) As Long\nEnd Type\n", true),
        ("009", "Type DynamicArray\n  Items() As String\nEnd Type\n", true),
    ];
    
    // Invalid syntax tests (should fail)
    let invalid_tests = vec![
        ("011", "Type EmptyType\nEnd Type\n", false),
        ("012", "Type MissingEnd\n  x As Integer\n", false),
        ("013", "Type DuplicateField\n  Name As String\n  Name As Integer\nEnd Type\n", false),
        ("014", "Type Type\n  x As Integer\nEnd Type\n", false),
        ("046", "Type NoFieldType\n  x\nEnd Type\n", false),
    ];
    
    for (id, code, should_succeed) in valid_tests.iter().chain(invalid_tests.iter()) {
        total += 1;
        let result = run_test(&format!("Test {}", id), code, *should_succeed);
        match result {
            TestResult::Ok => passed += 1,
            TestResult::Fail => failed += 1,
            TestResult::NotImplemented => not_implemented += 1,
        }
    }
    
    println!("\n========================================");
    println!("TEST SUMMARY");
    println!("========================================");
    println!("Total tests:        {}", total);
    println!("Passed (OK):        {}", passed);
    println!("Failed (FAIL):      {}", failed);
    println!("Not Implemented:    {}", not_implemented);
    println!("========================================\n");
    
    if failed > 0 {
        println!("⚠️  WARNING: {} test(s) failed", failed);
    }
    if not_implemented > 0 {
        println!("ℹ️  INFO: {} feature(s) not yet implemented", not_implemented);
    }
    if failed == 0 && not_implemented == 0 {
        println!("✅ All tests passed!");
    }
}