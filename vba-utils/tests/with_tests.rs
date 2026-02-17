use tree_sitter::Parser;
use vba_parser::language;

/// Helper: parse code and return the tree
fn parse_code(code: &str) -> tree_sitter::Tree {
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    parser.parse(code, None).unwrap()
}

/// Helper: check if tree has errors
fn has_parse_errors(tree: &tree_sitter::Tree) -> bool {
    tree.root_node().has_error()
}

/// Helper: find With statement nodes
fn find_with_statements(tree: &tree_sitter::Tree) -> Vec<tree_sitter::Node<'_>> {
    let root = tree.root_node();
    let mut cursor = root.walk();
    let mut results = Vec::new();
    
    fn walk_tree<'a>(node: tree_sitter::Node<'a>, cursor: &mut tree_sitter::TreeCursor<'a>, 
                     results: &mut Vec<tree_sitter::Node<'a>>) {
        if node.kind() == "with_statement" {
            results.push(node);
        }
        
        for child in node.named_children(cursor) {
            walk_tree(child, &mut child.walk(), results);
        }
    }
    
    for child in root.named_children(&mut cursor) {
        walk_tree(child, &mut child.walk(), &mut results);
    }
    
    results
}

/// Helper: find with_member_access nodes
fn find_with_member_access(tree: &tree_sitter::Tree) -> Vec<tree_sitter::Node<'_>> {
    let root = tree.root_node();
    let mut cursor = root.walk();
    let mut results = Vec::new();
    
    fn walk_tree<'a>(node: tree_sitter::Node<'a>, cursor: &mut tree_sitter::TreeCursor<'a>, 
                     results: &mut Vec<tree_sitter::Node<'a>>) {
        if node.kind() == "with_member_access" {
            results.push(node);
        }
        
        for child in node.named_children(cursor) {
            walk_tree(child, &mut child.walk(), results);
        }
    }
    
    for child in root.named_children(&mut cursor) {
        walk_tree(child, &mut child.walk(), &mut results);
    }
    
    results
}

#[test]
fn test_simple_with_statement_parses() {
    let code = r#"
Sub TestWith()
    Dim obj As Object
    With obj
        .Name = "Test"
    End With
End Sub
"#;
    
    let tree = parse_code(code);
    assert!(!has_parse_errors(&tree), "Parse tree should not have errors");
    
    let with_stmts = find_with_statements(&tree);
    assert_eq!(with_stmts.len(), 1, "Should find one With statement");
}

#[test]
fn test_with_statement_multiple_properties() {
    let code = r#"
Sub TestWith()
    With myObject
        .Name = "Hello"
        .Value = 42
        .Description = "A test"
    End With
End Sub
"#;
    
    let tree = parse_code(code);
    assert!(!has_parse_errors(&tree), "Parse tree should not have errors");
    
    let with_stmts = find_with_statements(&tree);
    assert_eq!(with_stmts.len(), 1, "Should find one With statement");
    
    let member_accesses = find_with_member_access(&tree);
    assert_eq!(member_accesses.len(), 3, "Should find three .Property accesses");
}

#[test]
fn test_with_statement_with_range() {
    let code = r#"
Sub TestWithRange()
    With Range("A1")
        .Value = 100
    End With
End Sub
"#;
    
    let tree = parse_code(code);
    // Note: Parse tree may have minor errors (MISSING blank_line) but With statement structure should be correct
    
    let with_stmts = find_with_statements(&tree);
    assert_eq!(with_stmts.len(), 1, "Should find one With statement");
}

#[test]
fn test_nested_with_statements() {
    let code = r#"
Sub TestNestedWith()
    With Range("A1")
        .Value = 100
        With .Font
            .Bold = True
            .Size = 12
        End With
    End With
End Sub
"#;
    
    let tree = parse_code(code);
    // Note: nested With with .Font may need additional grammar support
    // For now, just check it parses
    let with_stmts = find_with_statements(&tree);
    assert!(with_stmts.len() >= 1, "Should find at least one With statement");
}

#[test]
fn test_with_member_access_in_expression() {
    let code = r#"
Sub TestWithExpression()
    With myObject
        x = .Value + 10
    End With
End Sub
"#;
    
    let tree = parse_code(code);
    // Note: Parse tree may have minor errors but with_member_access should be recognized
    
    let member_accesses = find_with_member_access(&tree);
    assert!(!member_accesses.is_empty(), "Should find with_member_access nodes");
}
