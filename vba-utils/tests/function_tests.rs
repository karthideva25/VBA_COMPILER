use vba_utils::{walk_parse_tree, VbaAstNode};
use tree_sitter::Parser;
use vba_parser::language as vba_language;

fn parse_to_ast(code: &str) -> Vec<VbaAstNode> {
    let mut parser = Parser::new();
    parser.set_language(vba_language()).expect("Failed to load VBA language");
    let tree = parser.parse(code, None).expect("Parsing failed");
    let root_node = tree.root_node();
    let mut cursor = root_node.walk();
    walk_parse_tree(&mut cursor, code)
}

#[test]
fn test_simple_function_parsing() {
    let code = r#"
        Function AddNumbers(a As Integer, b As Integer) As Integer
            AddNumbers = a + b
        End Function
    "#;

    let ast = parse_to_ast(code);
    
    // Filter to find function nodes using if-let
    let functions: Vec<_> = ast.iter()
        .filter(|n| {
            if let VbaAstNode::Function { .. } = n {
                true
            } else {
                false
            }
        })
        .collect();
    
    assert_eq!(functions.len(), 1, "Expected 1 function, found {}", functions.len());

    match &functions[0] {
        VbaAstNode::Function { name, params, return_type, .. } => {
            assert_eq!(name, "AddNumbers");
            assert_eq!(params.len(), 2);
            assert_eq!(return_type.as_deref(), Some("Integer"));
        }
        _ => panic!("Expected a Function node"),
    }
}

#[test]
fn test_function_no_params() {
    let code = r#"
        Function GetValue() As String
            GetValue = "Hello"
        End Function
    "#;

    let ast = parse_to_ast(code);
    let functions: Vec<_> = ast.iter()
        .filter(|n| matches!(n, VbaAstNode::Function { .. }))
        .collect();
    
    assert_eq!(functions.len(), 1);

    match &functions[0] {
        VbaAstNode::Function { name, params, return_type, .. } => {
            assert_eq!(name, "GetValue");
            assert_eq!(params.len(), 0);
            assert_eq!(return_type.as_deref(), Some("String"));
        }
        _ => panic!("Expected a Function node"),
    }
}

#[test]
fn test_function_with_byval() {
    let code = r#"
        Function DoubleIt(ByVal x As Integer) As Integer
            DoubleIt = x * 2
        End Function
    "#;

    let ast = parse_to_ast(code);
    let functions: Vec<_> = ast.iter()
        .filter(|n| matches!(n, VbaAstNode::Function { .. }))
        .collect();
    
    assert_eq!(functions.len(), 1);

    match &functions[0] {
        VbaAstNode::Function { name, params, .. } => {
            assert_eq!(name, "DoubleIt");
            assert_eq!(params.len(), 1);
            // Check if ByVal is preserved in the Parameter struct
            assert_eq!(params[0].name, "x");
            assert_eq!(params[0].passing, vba_utils::ast::ParameterPassing::ByVal);
        }
        _ => panic!("Expected a Function node"),
    }
}

#[test]
fn test_function_with_optional() {
    let code = r#"
        Function Greet(Optional name As String = "World") As String
            Greet = "Hello, " & name
        End Function
    "#;

    let ast = parse_to_ast(code);
    let functions: Vec<_> = ast.iter()
        .filter(|n| matches!(n, VbaAstNode::Function { .. }))
        .collect();
    
    assert_eq!(functions.len(), 1);

    match &functions[0] {
        VbaAstNode::Function { name, params, .. } => {
            assert_eq!(name, "Greet");
            assert_eq!(params.len(), 1);
            assert!(params[0].optional);
        }
        _ => panic!("Expected a Function node"),
    }
}

#[test]
fn test_property_get() {
    let code = r#"
        Property Get Name() As String
            Name = mName
        End Property
    "#;

    let ast = parse_to_ast(code);
    let properties: Vec<_> = ast.iter()
        .filter(|n| matches!(n, VbaAstNode::PropertyGet { .. }))
        .collect();
    
    assert_eq!(properties.len(), 1, "Expected 1 property get, found {}", properties.len());

    match &properties[0] {
        VbaAstNode::PropertyGet { name, return_type, .. } => {
            assert_eq!(name, "Name");
            assert_eq!(return_type.as_deref(), Some("String"));
        }
        _ => panic!("Expected a PropertyGet node"),
    }
}

#[test]
fn test_property_let() {
    let code = r#"
        Property Let Name(value As String)
            mName = value
        End Property
    "#;

    let ast = parse_to_ast(code);
    let properties: Vec<_> = ast.iter()
        .filter(|n| matches!(n, VbaAstNode::PropertyLet { .. }))
        .collect();
    
    assert_eq!(properties.len(), 1);

    match &properties[0] {
        VbaAstNode::PropertyLet { name, params, .. } => {
            assert_eq!(name, "Name");
            assert_eq!(params.len(), 1);
        }
        _ => panic!("Expected a PropertyLet node"),
    }
}

#[test]
fn test_property_set() {
    let code = r#"
        Property Set MyObject(obj As Object)
            Set mObject = obj
        End Property
    "#;

    let ast = parse_to_ast(code);
    let properties: Vec<_> = ast.iter()
        .filter(|n| matches!(n, VbaAstNode::PropertySet { .. }))
        .collect();
    
    assert_eq!(properties.len(), 1);

    match &properties[0] {
        VbaAstNode::PropertySet { name, params, .. } => {
            assert_eq!(name, "MyObject");
            assert_eq!(params.len(), 1);
        }
        _ => panic!("Expected a PropertySet node"),
    }
}
