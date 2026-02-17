use vba_utils::{walk_parse_tree, VbaAstNode};
use tree_sitter::{Parser};
use vba_parser::language as vba_language;

fn parse_to_ast(code: &str) -> Vec<VbaAstNode> {
    let mut parser = Parser::new();
    parser.set_language(vba_language()).expect("Failed to load VBA language");
    let tree = parser.parse(code, None).expect("Parsing failed");
    let root_node = tree.root_node();
    let mut cursor = root_node.walk();
    let all = walk_parse_tree(&mut cursor, code);

    // Keep only subroutine nodes for tests that expect only subs
    all.into_iter()
        .filter(|n| matches!(n, VbaAstNode::Subroutine { .. }))
        .collect()
}

#[test]
fn test_single_subroutine_parsing() {
    let code = r#"
        Sub HelloWorld()
            MsgBox "Hello!"
        End Sub
    "#;

    let ast = parse_to_ast(code);
    assert_eq!(ast.len(), 1);

    match &ast[0] {
        VbaAstNode::Subroutine { name, body, .. } => {
            assert_eq!(name, "HelloWorld");
            assert!(!body.is_empty());
        }
        _ => panic!("Expected a Subroutine node"),
    }
}

#[test]
fn test_empty_subroutine() {
    let code = r#"
        Sub EmptySub()
        End Sub
    "#;

    let ast = parse_to_ast(code);
    assert_eq!(ast.len(), 1);

    match &ast[0] {
        VbaAstNode::Subroutine { name, body, .. } => {
            assert_eq!(name, "EmptySub");
            assert_eq!(body.len(), 0);
        }
        _ => panic!("Expected a Subroutine node"),
    }
}

#[test]
fn test_multiple_subroutines() {
    let code = r#"
        Sub First()
            MsgBox "First"
        End Sub

        Sub Second()
            MsgBox "Second"
        End Sub
    "#;

    let ast = parse_to_ast(code);
    assert_eq!(ast.len(), 2);

    match &ast[0] {
        VbaAstNode::Subroutine { name, .. } => assert_eq!(name, "First"),
        _ => panic!("Expected First Subroutine"),
    }

    match &ast[1] {
        VbaAstNode::Subroutine { name, .. } => assert_eq!(name, "Second"),
        _ => panic!("Expected Second Subroutine"),
    }
}
