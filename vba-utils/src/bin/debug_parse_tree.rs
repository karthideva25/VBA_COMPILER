use vba_utils::print_parse_tree;
use tree_sitter::Parser;
use vba_parser::language;

fn main() {
    // Test With Range statement
    let code_with = r#"Sub TestWithRange()
    With Range("A1")
        .Value = 100
    End With
End Sub"#;

    println!("=== With Range Test ===");
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code_with, None).unwrap();
    println!("Code:\n{}\n", code_with);
    println!("Has errors: {}", tree.root_node().has_error());
    println!("\nS-expression:\n{}", tree.root_node().to_sexp());
    
    println!("\n\n=== With Expression Test ===");
    let code_expr = r#"Sub TestWithExpression()
    x = 10
End Sub
"#;  // Assignment outside With
    let tree2 = parser.parse(code_expr, None).unwrap();
    println!("Code:\n{}\n", code_expr);
    println!("Has errors: {}", tree2.root_node().has_error());
    println!("\nS-expression:\n{}", tree2.root_node().to_sexp());

    println!("\n\n=== Single Subroutine Parse Tree ===");
    let code = r#"
        Sub HelloWorld()
            MsgBox "Hello!"
        End Sub
    "#;
    print_parse_tree(code);
}