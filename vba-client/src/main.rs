// use tree_sitter::Parser;
// use vba_parser::language;

// fn main() {
//     let mut parser = Parser::new();
//     parser.set_language(language()).unwrap();

//     let tree = parser.parse("Sub Hello()\nMsgBox \"Hi\"\nEnd Sub", None).unwrap();
//     println!("{}", tree.root_node().to_sexp());
// }

// vba-client/src/main.rs
use tree_sitter::{Parser};
use vba_parser::language as tree_sitter_vba;
use vba_utils::{print_parse_tree, execute_ast, Context, Program};
use vba_utils::ast::Statement;
fn main() {
    let vba_code = r#"
    Sub AutoOpen()   
       Call Sometest
       Call UseEmployeeType
        
    End Sub
    Type Employee
        FirstName As String
        LastName As String
    End Type
    Private Enum enu_SelectionCriteria
        Education = 0
        Experience = 1
        Certification = 2
    End Enum
    Sub UseEmployeeType()
        ' Declare a variable of the "Employee" type
        Dim Emp1 As Employee
        Emp1.FirstName = "John"
        Emp1.LastName = "Doe"
        MsgBox  Emp1.FirstName 
        MsgBox enu_SelectionCriteria.Certification
    End Sub
    Sub Sometest()
        Dim i As Integer
        Dim result As Double
        Dim divisor As Integer
        
        ' Using On Error Resume Next inside the For loop to skip errors on individual iterations
        For i = 1 To 5
            On Error GoTo ErrHandler
            
            ' Intentionally cause an error when divisor is zero
            divisor = 3 - i * 2 ' When i=2, divisor = -1, i=3 divisor = -3, i=4 divisor = -5, no zero yet; Let's fix it to cause zero at some point:
            ' Let's adjust so divisor will be zero when i=2
            divisor = 2 - i
            
            result = 10 / divisor
            
            If Err.Number <> 0 Then
                MsgBox "Error at iteration " & i & ": " & Err.Description
                Err.Clear
            Else
                MsgBox "Iteration " & i & ": result = " & result
            End If
        Next i
        ErrHandler:
            MsgBox "Error occurred!"
        
    End Sub
    Sub GradeCheck()
        Dim score As Integer
        score = 85
        If score >= 90 Then
            MsgBox "Grade: A"
        ElseIf score >= 80 Then
            MsgBox "Grade: B"
        Else
            MsgBox "Grade: C or lower"
        End If
    
    On Error GoTo ErrHandler
        Dim i As Integer, txt As String, val As Double, flag As Boolean
        val = 10
        flag = True
        For i = 1 To 3
            If flag Then
                txt = "Iteration " & i & ": " & val * i
                MsgBox txt
            Else
                MsgBox "Flag off"
            End If
        Next i
        Exit Sub
    ErrHandler:
        MsgBox "Error occurred!"
    End Sub
    Sub GotoStatementDemo() 
        Dim Number, MyString 
        Number = 1 ' Initialize variable. 
        ' Evaluate Number and branch to appropriate label. 
        If Number = 1 Then GoTo Line1 Else GoTo Line2 
        
        Line1: 
        MyString = "Number equals 1" 
        GoTo LastLine ' Go to LastLine. 
        Line2: 
        ' The following statement never gets executed. 
        MyString = "Number equals 2" 
        LastLine: 
        MsgBox MyString ' Print "Number equals 1" in 
        ' the Immediate window. 
    End Sub



    "#;

    // Set up parser
    let mut parser = Parser::new();
    parser.set_language(tree_sitter_vba()).expect("Failed to load grammar");

    // Parse input
    let tree = parser.parse(vba_code, None).expect("Parsing failed");
    let root_node = tree.root_node();

    // println!("\n🔍 Tree-sitter Parse Tree:");
    print_parse_tree(vba_code);  // <-- You already have a utility for this!

    // Build AST from the parse tree
    use vba_utils::ast::build_ast;
    let program = build_ast(root_node, vba_code);

        // DUMP THE WHOLE AST
        dbg!(&program);

    let mut ctx = Context::default();
    println!("\n🧠 Interpreting AST:");
    execute_ast(&program, &mut ctx).unwrap();


    // --- choose an entrypoint ---
    let mut entry: Option<String> = None;

    // prefer conventional names
    for candidate in ["Main", "main", "Auto_Open", "AutoOpen", "somemacro"] {
        if ctx.subs.contains_key(candidate) {
            entry = Some(candidate.to_string());
            break;
        }
    }

    // if there is exactly one sub, run it
    if entry.is_none() && ctx.subs.len() == 1 {
        entry = ctx.subs.keys().next().cloned();
    }

    // --- run it ---
    if let Some(name) = entry {
        println!("\n▶️ Running entrypoint: {}", name);
    
        // leverage the list driver used by execute_ast
        let call = vba_utils::ast::Statement::Call { function: name, args: vec![] };
        vba_utils::interpreter::execute_statement_list(std::slice::from_ref(&call), &mut ctx);
    } else {
        eprintln!("(No entrypoint to run. Defined subs: {:?})",
            ctx.subs.keys().collect::<Vec<_>>());
    }

}
//Sub somemacro()

// Dim j As Integer
// Dim h As Integer
// Dim a as Range
// set a=Range("a1")
// j = 10
// h = 45
// variable = "variable"
// variable = variable + j
// MsgBox (variable )
// Range("A1").Value
// call Integer
// call Range
// If j<100 Then
//     variable = "ten"
// Else 
//     variable = "not ten"
// End If

// End Sub


//     Sub HelloWorld()
//         MsgBox "Hello, World!"
//         Dim j As Integer
//         Dim h As Integer
//         j = 10
//         h = 45
//         If j Then
//             variable = "ten"
//         Else If
//             variable = "not ten"
//         End If

//     End Sub

// Sub B()
// A "john"
// End Sub
// Sub A(name)
// MsgBox "Hello, World!"
// Dim const as Variable
// Dim b as Const 
// Dim a as Variant
// Dim h as Integer
// End Sub

// Range("A1").Value = Cell
// MsgBox ("variable" )
// If j<100 Then
//     MsgBox "ten"
// Else 
//     MsgBox "not ten"
// End If


// Sub somemacro()

//         Dim j As Integer
//         Dim h As Integer
//         Dim m As Integer
//         m= 79
//         j = 10+ m +1000
//         h = vbRed
//         If j<100 Then
//              MsgBox "ten" & j
//         Else 
//              MsgBox "not ten"
//         End If
//         Dim i As Integer
//         i=1
//         For i To 5
//             MsgBox "Value of i = " & i
//         Next i
//         j=78
//         MsgBox h
//         Msgbox " I am a msgbox msg &j"
//         Msgbox " I am a msgbox msg" & j
//         MsgBox ("red " & vbRed)
       
//     End Sub

// Sub somemacro()
// Dim i As Integer
// For i = 1 To 5
// If i = 3 Then
//     Msgbox "Inside if" & i
//     Msgbox "Inside if  i should be 3 : " & i
    
//     Exit For
//     Msgbox "Inside if for exited" & i
// End If 
// Msgbox "count first " &i
// Next i
// For i = 1 To 5
// If i = 3 Then
//     GoTo AfterFor
// End If
// MsgBox "count" & i
// Next i
// AfterFor:
// MsgBox "done"


// MsgBox "I am some" & vbCrLf & "vbCrLf"

// End Sub