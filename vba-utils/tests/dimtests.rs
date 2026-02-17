use tree_sitter::Parser;
use vba_parser::language;

/// Helper: find the first "subroutine" node in the parse tree
fn find_first_subroutine(code: &str) -> tree_sitter::Tree {
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    tree
}

/// Helper: get the first subroutine node from a tree
fn get_first_subroutine_node(tree: &tree_sitter::Tree) -> tree_sitter::Node<'_> {
    let root = tree.root_node();
    let mut cursor = root.walk();
    
    let children: Vec<_> = root.named_children(&mut cursor).collect();
    
    // Try to find subroutine in statement children
    for child in &children {
        if child.kind() == "statement" {
            let mut stmt_cursor = child.walk();
            let stmt_children: Vec<_> = child.named_children(&mut stmt_cursor).collect();
            for stmt_child in stmt_children {
                if stmt_child.kind() == "subroutine" || stmt_child.kind() == "subroutine_definition" {
                    return stmt_child;
                }
            }
        }
    }
    
    // Fallback: look for direct subroutine children
    children.into_iter()
        .find(|n| n.kind() == "subroutine" || n.kind() == "subroutine_definition")
        .expect("Expected a top-level subroutine")
}

#[test]
fn parse_simple_dim_statement() {
    let code = "Sub Dims()\n    Dim j As Integer\nEnd Sub\n";
    let tree = find_first_subroutine(code);
    let sub = get_first_subroutine_node(&tree);
    let mut cursor = sub.walk();
    
    // Collect all statement nodes under the subroutine that contain Dim
    let dims: Vec<_> = sub.named_children(&mut cursor)
        .filter(|n| {
            n.kind() == "statement" && 
            n.utf8_text(code.as_bytes()).unwrap_or("").trim_start().starts_with("Dim")
        })
        .collect();
    assert_eq!(dims.len(), 1, "Expected exactly one Dim statement");
    let text = dims[0].utf8_text(code.as_bytes()).unwrap();
    assert_eq!(text, "Dim j As Integer\n");
}

#[test]
fn parse_dim_multiple_variables() {
    let code = "Sub Dims()\n    Dim a, b As String\nEnd Sub\n";
    let tree = find_first_subroutine(code);
    let sub = get_first_subroutine_node(&tree);
    let mut cursor = sub.walk();
    let dims: Vec<_> = sub.named_children(&mut cursor)
        .filter(|n| {
            n.kind() == "statement" && 
            n.utf8_text(code.as_bytes()).unwrap_or("").trim_start().starts_with("Dim")
        })
        .collect();
    assert_eq!(dims.len(), 1);
    let text = dims[0].utf8_text(code.as_bytes()).unwrap();
    assert_eq!(text, "Dim a, b As String\n");
}

#[test]
fn parse_dim_without_type() {
    let code = "Sub Dims()\n    Dim x\nEnd Sub\n";
    let tree = find_first_subroutine(code);
    let sub = get_first_subroutine_node(&tree);
    let mut cursor = sub.walk();
    let dims: Vec<_> = sub.named_children(&mut cursor)
        .filter(|n| {
            n.kind() == "statement" && 
            n.utf8_text(code.as_bytes()).unwrap_or("").trim_start().starts_with("Dim")
        })
        .collect();
    assert_eq!(dims.len(), 1);
    let text = dims[0].utf8_text(code.as_bytes()).unwrap();
    assert_eq!(text, "Dim x\n");
}

#[test]
fn parse_dim_multiple_as_types() {
    let code = "Sub Dims()\n    Dim i As Integer, s As String\nEnd Sub\n";
    let tree = find_first_subroutine(code);
    let sub = get_first_subroutine_node(&tree);
    let mut cursor = sub.walk();
    let dims: Vec<_> = sub.named_children(&mut cursor)
        .filter(|n| {
            n.kind() == "statement" && 
            n.utf8_text(code.as_bytes()).unwrap_or("").trim_start().starts_with("Dim")
        })
        .collect();
    assert_eq!(dims.len(), 1);
    let text = dims[0].utf8_text(code.as_bytes()).unwrap();
    assert_eq!(text, "Dim i As Integer, s As String\n");
}

#[test]
fn parse_multiple_dim_statements() {
    let code = "Sub Dims()\n    Dim x As Boolean\n    Dim y\n    Dim a, b As Variant\nEnd Sub\n";
    let tree = find_first_subroutine(code);
    let sub = get_first_subroutine_node(&tree);
    let mut cursor = sub.walk();
    let dims: Vec<_> = sub.named_children(&mut cursor)
        .filter(|n| {
            n.kind() == "statement" && 
            n.utf8_text(code.as_bytes()).unwrap_or("").trim_start().starts_with("Dim")
        })
        .collect();
    assert_eq!(dims.len(), 3);
    let texts: Vec<_> = dims.iter()
        .map(|n| n.utf8_text(code.as_bytes()).unwrap())
        .collect();
    assert_eq!(texts, vec![
        "Dim x As Boolean\n",
        "Dim y\n",
        "Dim a, b As Variant\n",
    ]);
}
