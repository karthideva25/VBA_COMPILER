use tree_sitter::Parser;
use vba_parser::language;
use vba_utils::{walk_parse_tree, VbaAstNode};

/// Extracts all subroutine AST nodes (name and body)
fn extract_all_subs(ast: Vec<VbaAstNode>) -> Vec<(String, Vec<VbaAstNode>)> {
    ast.into_iter()
        .filter_map(|n| {
            if let VbaAstNode::Subroutine { name, body, params: _ } = n {
                Some((name, body))
            } else {
                None
            }
        })
        .collect()
}

#[test]
fn parse_simple_subroutine() {
    let code = "Sub Foo()
End Sub
";
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    let mut cursor = tree.root_node().walk();
    let ast = walk_parse_tree(&mut cursor, code);

    let subs = extract_all_subs(ast);
    assert_eq!(subs.len(), 1, "Should find exactly one subroutine");
    assert_eq!(subs[0].0, "Foo");
}

#[test]
fn parse_subroutine_with_inner_call() {
    let code = "Sub Outer()
    Inner()
End Sub
";
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    let mut cursor = tree.root_node().walk();
    let ast = walk_parse_tree(&mut cursor, code);

    let subs = extract_all_subs(ast);
    assert_eq!(subs.len(), 1, "Should find one subroutine");
    assert_eq!(subs[0].0, "Outer");
}

#[test]
fn parse_sub_with_multiple_calls() {
    let code = "Sub Caller()
    One()
    Two()
End Sub
";
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    let mut cursor = tree.root_node().walk();
    let ast = walk_parse_tree(&mut cursor, code);

    let subs = extract_all_subs(ast);
    assert_eq!(subs.len(), 1);
    assert_eq!(subs[0].0, "Caller");
}

#[test]
fn parse_sub_with_param_and_nonparam_calls() {
    let code = "Sub Params()
    NoParam()
    Param(42)
End Sub
";
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    let mut cursor = tree.root_node().walk();
    let ast = walk_parse_tree(&mut cursor, code);

    let subs = extract_all_subs(ast);
    assert_eq!(subs.len(), 1);
    assert_eq!(subs[0].0, "Params");
}

#[test]
fn parse_sub_with_msgbox_print_name() {
    let code = "Sub ShowName()
    MsgBox \"ShowName\"
End Sub
";
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    let mut cursor = tree.root_node().walk();
    let ast = walk_parse_tree(&mut cursor, code);

    let subs = extract_all_subs(ast);
    assert_eq!(subs.len(), 1);
    assert_eq!(subs[0].0, "ShowName");
}

#[test]
fn parse_two_subroutines() {
    let code = r#"
Sub A()
End Sub

Sub B()
End Sub
"#;
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    let mut cursor = tree.root_node().walk();
    let ast = walk_parse_tree(&mut cursor, code);

    let subs = extract_all_subs(ast);
    assert_eq!(subs.len(), 2);
    assert_eq!(subs[0].0, "A");
    assert_eq!(subs[1].0, "B");
}

#[test]
fn parse_custom_subroutines() {
    let code = r#"
Sub A(name)
End Sub

Sub B()
   a "john"
End Sub
"#;
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    let mut cursor = tree.root_node().walk();
    let ast = walk_parse_tree(&mut cursor, code);

    let subs = extract_all_subs(ast);
    assert_eq!(subs.len(), 2);
    assert_eq!(subs[0].0, "A");
    assert_eq!(subs[1].0, "B");
}

#[test]
fn parse_sub_with_parameters() {
    let code = "Sub WithParams(a, b, c)
End Sub
";
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    let root = tree.root_node();
    let mut cursor = root.walk();

    // Extract parameter identifiers from the first top-level subroutine
    let mut params = Vec::new();
    for node in root.named_children(&mut cursor) {
        if node.kind() == "subroutine_definition" {
            let mut sub_cur = node.walk();
            for pl in node.named_children(&mut sub_cur).filter(|n| n.kind() == "parameter_list") {
                let mut pcur = pl.walk();
                for child in pl.named_children(&mut pcur).filter(|n| n.kind() == "identifier") {
                    let s = child.utf8_text(code.as_bytes()).unwrap();
                    params.push(s.to_string());
                }
            }
            break;
        }
    }
    assert_eq!(params, vec!["a", "b", "c"]);
}

#[test]
fn reject_inline_sub_without_newline() {
    let code = "Sub WithParams(a, b, c) End Sub";
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    let mut cursor = tree.root_node().walk();
    let ast = walk_parse_tree(&mut cursor, code);
    let subs = extract_all_subs(ast);
    assert!(subs.is_empty(), "Inline Sub on one line should not be recognized");
}
#[test]
fn parse_bare_call_with_space() {
    let code = r#"Sub B()
    a "john"
End Sub
"#;

    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();

    // Build high-level AST nodes (subs, etc.)
    let mut cursor = tree.root_node().walk();
    let ast = walk_parse_tree(&mut cursor, code);

    // Keep only subroutines (preserves order)
    let subs = extract_all_subs(ast);

    // Find the target sub
    let (name, body) = subs
        .into_iter()
        .find(|(n, _)| n.eq_ignore_ascii_case("B"))
        .expect("expected sub B");

    assert_eq!(name, "B");

    // Search inside the sub body for a bare Call to `a`
    let mut found_call = false;
    for node in body {
        if let VbaAstNode::Call { function, args } = node {
            if function.eq_ignore_ascii_case("a") {
                // Expect exactly one argument, a string literal "john"
                assert_eq!(args.len(), 1, "call should have one argument");
                match &args[0] {
                    VbaAstNode::ExpressionString(s)         // if your walker wraps string as a specific node
                    | VbaAstNode::String(s)                 // or this variant
                    => assert_eq!(s, "john"),
                    VbaAstNode::Expression { kind, .. }     // or if expressions are wrapped
                    if kind == "String" => { /* optionally validate inner text */ }
                    other => panic!("expected string literal arg, got {:?}", other),
                }
                found_call = true;
                break;
            }
        }
    }

    assert!(found_call, "Expected to find a bare call to 'a' with a single string arg");
}
