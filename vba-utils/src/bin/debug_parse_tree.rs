use vba_utils::print_parse_tree;

fn main() {
    let code = r#"
        Sub HelloWorld()
            MsgBox "Hello!"
        End Sub
    "#;
    
    println!("=== Single Subroutine Parse Tree ===");
    print_parse_tree(code);
    
    let code2 = r#"
        Sub EmptySub()
        End Sub
    "#;
    
    println!("\n=== Empty Subroutine Parse Tree ===");
    print_parse_tree(code2);
    
    let code3 = r#"
        Sub First()
            MsgBox "First"
        End Sub

        Sub Second()
            MsgBox "Second"
        End Sub
    "#;
    
    println!("\n=== Multiple Subroutines Parse Tree ===");
    print_parse_tree(code3);
}