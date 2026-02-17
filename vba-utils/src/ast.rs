
use tree_sitter::Node;

/// A whole VBA program.
#[derive(Debug, Clone)]
pub struct Program {
    pub statements: Vec<Statement>,
}

/// All the statement kinds in your grammar.
#[derive(Debug,Clone)]
pub enum Statement {
    BlankLine,
    Comment(String),
    OptionExplicit,
    Subroutine {
        name: String,
        params: Vec<String>,
        body: Vec<Statement>,
    },
    Dim {
        names: Vec<(String, Option<String>)>,
    },
    ReDim {  // ADD THIS
        preserve: bool,
        arrays: Vec<ReDimArray>,
    },
    Set {
        target: String,
        expr: Expression,
    },
    Assignment {
        lvalue: AssignmentTarget,
        rvalue: Expression,
    },
    MsgBox {
        expr: Expression,
    },
    GoTo {
        label: String,
    },
    If {
        condition: Expression,
        then_branch: Vec<Statement>,
        else_if: Vec<(Expression, Vec<Statement>)>,
        else_branch: Vec<Statement>,
    },
    For(ForStatement),
    DoWhile(DoWhileStatement),
    Exit(ExitType), 
    Enum {                              
        visibility: Option<String>,     
        name: String,                  
        members: Vec<EnumMember>,      
    },
    Type {                            
        visibility: Option<String>,     
        name: String,                  
        fields: Vec<TypeField>,
    },
    Label(String),
    Expression(Expression),
    OnError(OnErrorKind),
    Resume(ResumeKind),
    Call {
        function: String,
        args: Vec<Expression>,
    },
    
}

/// All the expression kinds in your grammar.
#[derive(Debug, Clone)]
pub enum Expression {
    Integer(i64),
    Byte(u8),
    Single(f32), 
    String(String),
    Identifier(String),
    Boolean(bool),
    Currency(f64),
    Date(chrono::NaiveDate), // Use chrono for dates
    Double(f64),       // ✅ add
    Decimal(f64),      // ✅ add (or use rust_decimal::Decimal if you want fixed precision)
    // Binary {
    //     left: Box<Expression>,
    //     operator: String,
    //     right: Box<Expression>,
    // },
    BinaryOp {
        left: Box<Expression>,
        op: String,
        right: Box<Expression>,
    },
    UnaryOp {
        op: String,
        expr: Box<Expression>,
    },
    FunctionCall {
        function: Box<Expression>,
        args: Vec<Expression>,
    },
    PropertyAccess {
        obj: Box<Expression>,
        property: String,
    },
    BuiltInConstant(String), 
}

#[derive(Debug, Clone)]
pub struct ForStatement {
    pub counter: String,              // Loop variable name (e.g., "i")
    pub start: Expression,            // Initial value expression
    pub end: Expression,              // End value expression  
    pub step: Option<Expression>,     // Optional step expression
    pub body: Vec<Statement>,         // Loop body statements
    pub next_counter: Option<String>, // Optional counter after Next (for validation)
}

#[derive(Debug, Clone)]
pub struct DoWhileStatement {
    pub condition: Option<Expression>,     // None for infinite Do...Loop
    pub condition_type: DoWhileConditionType,
    pub test_at_end: bool,                 // true = Do...Loop While, false = Do While...Loop
    pub body: Vec<Statement>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DoWhileConditionType {
    While,   // Continue while true
    Until,   // Continue until true (i.e., while false)
    Infinite,    // Infinite loop (Do...Loop)
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum OnErrorKind {
    ResumeNext,          // On Error Resume Next
    GoToLabel(String),   // On Error GoTo <label>
    GoToZero,            // On Error GoTo 0
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ResumeKind {
    Current,             // Resume
    Next,                // Resume Next
    Label(String),       // Resume <label>
}

#[derive(Debug, Clone, Default)]
pub struct ErrObject {
    pub number: i32,
    pub description: String,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ExitType {
    For,
    Do,
    While,
    Sub,
    Function,
    Property,
    Select,
}

// Add new struct for enum members
#[derive(Debug, Clone)]
pub struct EnumMember {
    pub name: String,                   // Member name
    pub value: Option<Expression>,      // Optional explicit value
}

#[derive(Debug, Clone)]
pub struct TypeField {
    pub name: String,                   // Field name
    pub field_type: String,             // Field type (Integer, String, custom type, etc.)
    pub dimensions: Option<Vec<ArrayDimension>>,  // For array fields
    pub string_length: Option<i64>,     // For fixed-length strings (String * 30)
}

#[derive(Debug, Clone)]
pub struct ArrayDimension {
    pub lower: Option<Expression>,      // Lower bound (optional)
    pub upper: Expression,              // Upper bound
}

#[derive(Debug, Clone)]
pub struct ReDimArray {
    pub name: String,
    pub dimensions: Vec<ArrayDimension>,
}

// Add this new enum:
#[derive(Debug, Clone)]
pub enum AssignmentTarget {
    Identifier(String),              // Simple: x = 5
    PropertyAccess {                 // Property: Emp1.FirstName = "John"
        object: String,
        property: String,
    },
    IndexedAccess {                  // ADD THIS - Array: arr(i) = 5
        array: String,
        indices: Vec<Expression>,
    },
    // Could add more complex targets later
}

impl std::fmt::Display for AssignmentTarget {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            AssignmentTarget::Identifier(name) => write!(f, "{}", name),
            AssignmentTarget::PropertyAccess { object, property } => write!(f, "{}.{}", object, property),
            AssignmentTarget::IndexedAccess { array, indices } => {
                write!(f, "{}({} indices)", array, indices.len())
            }
        }
    }
}


/// Build the top-level AST from the `source_file` node.
pub fn build_ast(root: Node, source: &str) -> Program {
    let mut stmts = Vec::new();
    let mut cursor = root.walk();
    for stmt_wr in root.named_children(&mut cursor) {
        if let Some(stmt) = build_statement(stmt_wr, source) {
            // Optionally filter out comments during execution
            match &stmt {
                Statement::Comment(_) => {
                    // You can choose to include or exclude comments from execution
                    // For now, let's include them for debugging
                    stmts.push(stmt);
                }
                _ => {
                    stmts.push(stmt);
                }
            }
        }
    }
    Program { statements: stmts }
}

/// Recursively build a Statement, unwrapping the generic `"statement"` wrappers.
fn build_statement(node: Node, source: &str) -> Option<Statement> {
    // eprintln!(
    //     "🔹 build_statement: kind = {:15} text = {:?}",
    //     node.kind(),
    //     node.utf8_text(source.as_bytes()).unwrap_or("")
    // );
    // 1) Unwrap the "statement" wrapper if present.
    if node.kind() == "statement" {
        let mut c = node.walk();
        for inner in node.named_children(&mut c) {
            return build_statement(inner, source);
        }
        return None;
    }

    match node.kind() {
        "blank_line" => Some(Statement::BlankLine),
        "comment" => {
            let comment_text = extract(source, node);
            // Remove the leading ' character
            let cleaned_comment = comment_text.strip_prefix('\'').unwrap_or(&comment_text);
            Some(Statement::Comment(cleaned_comment.trim().to_string()))
        }

        "subroutine" => {
            let name_node = node.child_by_field_name("name")?;
            let name = extract(source, name_node);
            eprintln!("🔨 Building Subroutine `{}`", name);

            // Extract parameters by iterating all named children
            let mut params = Vec::new();
            let mut cursor = node.walk();
            
            for child in node.named_children(&mut cursor) {
                if child.kind() == "parameter_list" {
                    eprintln!("  🔍 Found parameter_list node");
                    
                    for i in 0..child.child_count() {
                        if let Some(param_child) = child.child(i) {
                            if param_child.kind() == "identifier" {
                                let param_name = extract(source, param_child);
                                eprintln!("    ✅ Parameter: {}", param_name);
                                params.push(param_name);
                            }
                        }
                    }
                }
            }

            eprintln!("  ✅ Total parameters: {}", params.len());

            // Body
            let mut body = Vec::new();
            let mut bc = node.walk();
            for stmt_wrapper in node.named_children(&mut bc).filter(|n| n.kind() == "statement") {
                if let Some(stmt) = build_statement(stmt_wrapper, source) {
                    body.push(stmt);
                }
            }

            eprintln!("🔚 Subroutine `{}`: {} params, {} statements\n", name, params.len(), body.len());
            Some(Statement::Subroutine { name, params, body })
        }


        "dim_statement" => {
            let mut names = Vec::new();
            let mut dc = node.walk();

            let mut child_cursor = node.walk();
            let children: Vec<_> = node.named_children(&mut child_cursor).collect();

            // Iterate over children and detect (identifier, type) pairs
            let mut i = 0;
            while i < children.len() {
                let id = &children[i];
                if id.kind() == "identifier" {
                    let var = extract(source, *id);
                    let mut ty: Option<String> = None;

                    // Look ahead for a following type (primitive_type or identifier)
                    if i + 1 < children.len() {
                        let next = &children[i + 1];
                        if next.kind() == "primitive_type" || next.kind() == "identifier" {
                            ty = Some(extract(source, *next));
                            i += 1; // skip the type node
                        }
                    }

                    names.push((var, ty));
                }

                i += 1;
            }

            Some(Statement::Dim { names })
        }

        // In build_statement function, add this case (after "dim_statement"):
        "redim_statement" => {
            let mut preserve = false;
            let mut arrays = Vec::new();
            
            // Check for Preserve keyword
            let text = extract(source, node).to_lowercase();
            if text.contains("preserve") {
                preserve = true;
            }
            
            // Extract array redefinitions
            let mut cursor = node.walk();
            for child in node.named_children(&mut cursor) {
                if child.kind() == "identifier" {
                    let array_name = extract(source, child);
                    
                    // Look for dimensions after this identifier
                    let mut dimensions = Vec::new();
                    
                    // Find the next siblings that are array_dimension
                    let mut sibling_cursor = node.walk();
                    let mut found_this_id = false;
                    
                    for sibling in node.named_children(&mut sibling_cursor) {
                        if sibling.kind() == "identifier" && extract(source, sibling) == array_name {
                            found_this_id = true;
                            continue;
                        }
                        
                        if found_this_id && sibling.kind() == "array_dimension" {
                            if let Some(dim) = build_array_dimension(sibling, source) {
                                dimensions.push(dim);
                            }
                        } else if found_this_id && sibling.kind() == "identifier" {
                            // Hit next array name, stop collecting dimensions
                            break;
                        }
                    }
                    
                    if !dimensions.is_empty() {
                        arrays.push(ReDimArray {
                            name: array_name,
                            dimensions,
                        });
                    }
                }
            }
            
            eprintln!("✅ Built ReDim: preserve={}, arrays={}", preserve, arrays.len());
            
            Some(Statement::ReDim { preserve, arrays })
        }

        "set_statement" => {
            let mut sc = node.walk();
            let mut parts = node
                .children(&mut sc)
                .filter(|n| n.kind() == "identifier" || n.kind() == "expression");
            let target = extract(source, parts.next()?);
            let expr = build_expression(parts.next()?, source)?;
            Some(Statement::Set { target, expr })
        }
 
        "assignment_statement" => {
            let mut target: Option<AssignmentTarget> = None;
            let mut expr: Option<Expression> = None;
            
            let mut ac = node.walk();
            for child in node.named_children(&mut ac) {
                match child.kind() {
                    "lvalue" => {
                        // Extract identifier or property_access or indexed_access from lvalue node
                        let mut lvalue_cursor = child.walk();
                        for lvalue_child in child.named_children(&mut lvalue_cursor) {
                            match lvalue_child.kind() {
                                "identifier" => {
                                    let name = extract(source, lvalue_child);
                                    target = Some(AssignmentTarget::Identifier(name));
                                    break;
                                }
                                "property_access" => {
                                    let mut pc = lvalue_child.walk();
                                    let parts: Vec<_> = lvalue_child.named_children(&mut pc).collect();
                                    
                                    if parts.len() == 2 {
                                        let obj = extract(source, parts[0]);
                                        let prop = extract(source, parts[1]);
                                        eprintln!("🔍 Parsed property_access: object='{}', property='{}'", obj, prop);
                                        target = Some(AssignmentTarget::PropertyAccess {
                                            object: obj,
                                            property: prop,
                                        });
                                    } else {
                                        let full_text = extract(source, lvalue_child);
                                        eprintln!("⚠️ property_access has {} parts, using text fallback: '{}'", parts.len(), full_text);
                                        if let Some(dot_pos) = full_text.find('.') {
                                            let object = full_text[..dot_pos].to_string();
                                            let property = full_text[dot_pos + 1..].to_string();
                                            target = Some(AssignmentTarget::PropertyAccess { object, property });
                                        } else {
                                            target = Some(AssignmentTarget::Identifier(full_text));
                                        }
                                    }
                                    break;
                                }
                                "indexed_access" => {
                                    // Parse array(index) = value
                                    let mut array_name = None;
                                    let mut indices = Vec::new();
                                    
                                    let mut idx_cursor = lvalue_child.walk();
                                    for idx_child in lvalue_child.named_children(&mut idx_cursor) {
                                        match idx_child.kind() {
                                            "identifier" if array_name.is_none() => {
                                                array_name = Some(extract(source, idx_child));
                                            }
                                            "expression" => {
                                                if let Some(index_expr) = build_expression(idx_child, source) {
                                                    indices.push(index_expr);
                                                }
                                            }
                                            _ => {}
                                        }
                                    }
                                    
                                    if let Some(arr) = array_name {
                                        eprintln!("🔍 Parsed indexed_access lvalue: array='{}', indices={}", arr, indices.len());
                                        target = Some(AssignmentTarget::IndexedAccess {
                                            array: arr,
                                            indices,
                                        });
                                    }
                                    break;
                                }
                                _ => {}
                            }
                        }
                    }
                    "expression" => {
                        expr = build_expression(child, source);
                    }
                    "ERROR" | "=" => {
                        continue;
                    }
                    _ => {}
                }
            }
            
            // Fallback: try the old method if lvalue approach didn't work
            if target.is_none() || expr.is_none() {
                let mut ac2 = node.walk();
                let mut parts = node
                    .children(&mut ac2)
                    .filter(|n| n.kind() == "identifier" || n.kind() == "expression");
                if target.is_none() {
                    if let Some(id_node) = parts.next() {
                        let name = extract(source, id_node);
                        target = Some(AssignmentTarget::Identifier(name));
                    }
                }
                if expr.is_none() {
                    if let Some(expr_node) = parts.next() {
                        expr = build_expression(expr_node, source);
                    }
                }
            }
            
            if let (Some(target_val), Some(expression)) = (target.clone(), expr.clone()) {
                Some(Statement::Assignment { lvalue: target_val, rvalue: expression })
            } else {
                eprintln!("⚠️ Failed to build assignment statement - target: {:?}, expr: {:?}", &target, &expr);
                None
            }
        }

        "msgbox_statement" => {
            let mut expressions = Vec::new();
            
            let mut cursor = node.walk();
            for child in node.children(&mut cursor) {
                match child.kind() {
                    "expression" => {
                        // Direct expression child
                        if let Some(expr) = build_expression(child, source) {
                            expressions.push(expr);
                        }
                    }
                    "ERROR" => {
                        // Look inside ERROR node for expressions
                        let mut err_cursor = child.walk();
                        for err_child in child.children(&mut err_cursor) {
                            if err_child.kind() == "expression" {
                                if let Some(expr) = build_expression(err_child, source) {
                                    expressions.push(expr);
                                }
                            }
                        }
                    }
                    _ => {}
                }
            }
            
            // First expression is the message to display
            let message = expressions.into_iter().next();
            
            if let Some(msg_expr) = message {
                eprintln!("  ✅ built MsgBox with expr: {:?}", msg_expr);
                Some(Statement::MsgBox { expr: msg_expr })
            } else {
                eprintln!("  ⚠️ MsgBox has no expression");
                None
            }
        }

        "goto_statement" => {
            let mut gc = node.walk();
            let label = node
                .children(&mut gc)
                .find(|n| n.kind() == "identifier")
                .map(|n| extract(source, n))?;
            Some(Statement::GoTo { label })
        }

        "if_statement" => {
            // --- small helpers -------------------------------------------------------
            let get_between = |src: &str, a_end: usize, b_start: usize| -> String {
                if a_end >= b_start || b_start > src.len() { return String::new(); }
                src[a_end..b_start].to_string()  // Don't trim or lowercase yet
            };
            let has = |gap: &str, kw: &str| gap.to_ascii_lowercase().contains(kw);
            
            let mut has_newline_after_then = false;
            let mut has_end_if = false;
            let mut is_inline_form = false;
            
            // --- statement being built -----------------------------------------------
            let mut condition: Option<Expression> = None;
            let mut then_branch: Vec<Statement> = Vec::new();
            let mut else_if: Vec<(Expression, Vec<Statement>)> = Vec::new();
            let mut else_branch: Vec<Statement> = Vec::new();

            // Sections: before_then | then_body | elseif_condition | elseif_body | else_body
            let mut current_section = "before_then";
            let mut current_elseif_condition: Option<Expression> = None;
            let mut current_elseif_statements: Vec<Statement> = Vec::new();

            // iterate children with access to byte gaps
            let mut ic = node.walk();
            let children: Vec<Node> = node.children(&mut ic).collect();
            let nstart = node.start_byte();
            let nend   = node.end_byte();

            // First pass: detect if inline form and check for End If
            for (i, child) in children.iter().enumerate() {
                if child.kind() == "keyword_Then" && i + 1 < children.len() {
                    let next_start = children[i+1].start_byte();
                    let gap = get_between(source, child.end_byte(), next_start);
                    
                    // Check if there's a newline after Then
                    if gap.contains('\n') || gap.contains('\r') {
                        has_newline_after_then = true;
                    } else {
                        // Inline form: no newline after Then
                        is_inline_form = true;
                    }
                }
                
                if child.kind() == "keyword_End_If" {
                    has_end_if = true;
                }
            }

            // Validation
            if is_inline_form && has_end_if {
                eprintln!("⚠️ Warning: Inline If should not have End If at line {}", 
                        node.start_position().row + 1);
            }
            
            if has_newline_after_then && !has_end_if {
                eprintln!("⚠️ Warning: Block If missing End If at line {}", 
                        node.start_position().row + 1);
            }

            // Second pass: build the statement
            for (i, child) in children.iter().enumerate() {
                let prev_end   = if i == 0 { nstart } else { children[i-1].end_byte() };
                let next_start = if i + 1 < children.len() { children[i+1].start_byte() } else { nend };
                let gap_before = get_between(source, prev_end, child.start_byte());
                let gap_after  = get_between(source, child.end_byte(), next_start);

                // --- HANDLE EXPLICIT KEYWORD NODES FIRST ------------------------------
                match child.kind() {
                    "keyword_Then" => {
                        match current_section {
                            "before_then"      => current_section = "then_body",
                            "elseif_condition" => current_section = "elseif_body",
                            "elseif_body"      => { /* stay */ }
                            "else_body"        => { /* ignore stray THEN in else */ }
                            _                  => { /* no-op */ }
                        }
                        continue;
                    }

                    "keyword_ElseIf" => {
                        if let Some(cond) = current_elseif_condition.take() {
                            else_if.push((cond, std::mem::take(&mut current_elseif_statements)));
                        }
                        current_section = "elseif_condition";
                        continue;
                    }

                    "keyword_Else" => {
                        if let Some(cond) = current_elseif_condition.take() {
                            else_if.push((cond, std::mem::take(&mut current_elseif_statements)));
                        }
                        current_section = "else_body";
                        continue;
                    }

                    // Handle End If properly
                    "keyword_End_If" => {
                        if let Some(cond) = current_elseif_condition.take() {
                            else_if.push((cond, std::mem::take(&mut current_elseif_statements)));
                        }
                        break; // End of if statement
                    }

                    // Legacy handling for separate keyword_End
                    "keyword_End" => {
                        if has(&gap_after, "if") {
                            if let Some(cond) = current_elseif_condition.take() {
                                else_if.push((cond, std::mem::take(&mut current_elseif_statements)));
                            }
                            break;
                        }
                    }

                    _ => {}
                }

                // --- FALLBACK GAP SWITCHES ----
                if has(&gap_before, "end if") {
                    if let Some(cond) = current_elseif_condition.take() {
                        else_if.push((cond, std::mem::take(&mut current_elseif_statements)));
                    }
                    break;
                }
                if has(&gap_before, "else if") || has(&gap_before, "elseif") {
                    if let Some(cond) = current_elseif_condition.take() {
                        else_if.push((cond, std::mem::take(&mut current_elseif_statements)));
                    }
                    current_section = "elseif_condition";
                } else if has(&gap_before, "else") {
                    if let Some(cond) = current_elseif_condition.take() {
                        else_if.push((cond, std::mem::take(&mut current_elseif_statements)));
                    }
                    current_section = "else_body";
                } else if has(&gap_before, "then") && current_section == "before_then" {
                    current_section = "then_body";
                }

                // --- CONSUME NODE BY KIND --------------------------------------------
                match child.kind() {
                    "expression" => {
                        if current_section == "before_then" && condition.is_none() {
                            condition = build_expression(*child, source);
                            if has(&gap_after, "then") {
                                current_section = "then_body";
                            }
                        } else if current_section == "elseif_condition" {
                            current_elseif_condition = build_expression(*child, source);
                            current_section = "elseif_body";
                        } else {
                            if let Some(cond) = current_elseif_condition.take() {
                                else_if.push((cond, std::mem::take(&mut current_elseif_statements)));
                            }
                            current_elseif_condition = build_expression(*child, source);
                            current_section = "elseif_body";
                        }
                    }

                    "statement" => {
                        if let Some(stmt) = build_statement(*child, source) {
                            match current_section {
                                "then_body"   => then_branch.push(stmt),
                                "elseif_body" => current_elseif_statements.push(stmt),
                                "else_body"   => else_branch.push(stmt),
                                _             => { /* ignore stray statements */ }
                            }
                        }
                    }

                    _ => {}
                }
            }

            // flush any trailing ElseIf at end of block
            if let Some(cond) = current_elseif_condition.take() {
                else_if.push((cond, std::mem::take(&mut current_elseif_statements)));
            }

            // Debug output
            if is_inline_form {
                println!("  📄 Inline If: then={}, else={}", then_branch.len(), else_branch.len());
            } else {
                println!("  📦 Block If: then={}, elseif={}, else={}", 
                        then_branch.len(), else_if.len(), else_branch.len());
            }

            if let Some(cond) = condition {
                Some(Statement::If {
                    condition: cond,
                    then_branch,
                    else_if,
                    else_branch,
                })
            } else {
                eprintln!("Failed to build if statement - no condition found");
                None
            }
        }

        "for_statement" => {
            let mut counter = None;
            let mut start = None;
            let mut end = None;
            let mut step = None;
            let mut body = Vec::new();
            let mut next_counter = None;

            // Extract counter (loop variable)
            if let Some(counter_node) = node.child_by_field_name("counter") {
                counter = Some(extract(source, counter_node));
            }
            
            // Extract start expression
            if let Some(start_node) = node.child_by_field_name("start") {
                start = build_expression(start_node, source);
            }
            
            // Extract end expression
            if let Some(end_node) = node.child_by_field_name("end") {
                end = build_expression(end_node, source);
            }
            
            // Extract optional step expression
            if let Some(step_node) = node.child_by_field_name("step") {
                step = build_expression(step_node, source);
            }
            
            // Extract body statements - look for "statement" wrappers inside the for_statement
            let mut fc = node.walk();
            for child in node.named_children(&mut fc) {
                if child.kind() == "statement" {
                    if let Some(stmt) = build_statement(child, source) {
                        body.push(stmt);
                    }
                }
            }
            
            // Extract optional next counter (for validation)
            if let Some(nc) = node.child_by_field_name("next_counter") {
                next_counter = Some(extract(source, nc));
            }
            
            if let (Some(counter_name), Some(start_expr), Some(end_expr)) = (counter, start, end) {
                Some(Statement::For(ForStatement {
                    counter: counter_name,
                    start: start_expr,
                    end: end_expr,
                    step,
                    body,
                    next_counter,
                }))
            } else {
                eprintln!("Failed to build for statement - missing required components");
                None
            }
        }

        
        "do_while_statement" => {
            let mut condition: Option<Expression> = None;
            let mut condition_type = DoWhileConditionType::Infinite;
            let mut test_at_end = false;
            let mut body = Vec::new();
            
            // Extract condition if present
            if let Some(cond_node) = node.child_by_field_name("condition") {
                condition = build_expression(cond_node, source);
            }
            
            // Determine the loop type by examining the source text
            let loop_text = extract(source, node).to_lowercase();
            
            if loop_text.contains("do while") {
                condition_type = DoWhileConditionType::While;
                test_at_end = false;
            } else if loop_text.contains("do until") {
                condition_type = DoWhileConditionType::Until;
                test_at_end = false;
            } else if loop_text.contains("loop while") {
                condition_type = DoWhileConditionType::While;
                test_at_end = true;
            } else if loop_text.contains("loop until") {
                condition_type = DoWhileConditionType::Until;
                test_at_end = true;
            } else {
                // Plain Do...Loop (infinite)
               condition_type = DoWhileConditionType::Infinite;
            }
            
            // Extract body statements
            let mut cursor = node.walk();
            for child in node.named_children(&mut cursor) {
                if child.kind() == "statement" {
                    if let Some(stmt) = build_statement(child, source) {
                        body.push(stmt);
                    }
                }
            }
            
            eprintln!("✅ Built DoWhile loop: type={:?}, test_at_end={}, has_condition={}", 
                    condition_type, test_at_end, condition.is_some());
            
            Some(Statement::DoWhile(DoWhileStatement {
                condition,
                condition_type,
                test_at_end,
                body,
            }))
        }

        "exit_statement" => {
            // Preferred path: use the grammar field if present.
            if let Some(exit_type_node) = node.child_by_field_name("exit_type") {
                let exit_type_str = extract(source, exit_type_node).trim().to_string();
                if let Some(exit_type) = ExitType::from_str(&exit_type_str) {
                    return Some(Statement::Exit(exit_type));
                } else {
                    eprintln!("Unknown exit type (field): {}", exit_type_str);
                    // fall through to raw-text fallback
                }
            }

            // Fallback: parse raw node text like "Exit For", "Exit Do", … even if the field is missing.
            let raw = extract(source, node);
            // Grab the token after "Exit" (handles extra whitespace and a trailing ':' if present)
            if let Some(tok) = raw.split_whitespace().nth(1) {
                let cleaned = tok.trim_end_matches(':').trim();
                if let Some(exit_type) = ExitType::from_str(cleaned) {
                    return Some(Statement::Exit(exit_type));
                } else {
                    eprintln!("Unknown exit type (raw): {}", cleaned);
                }
            } else {
                eprintln!("Missing exit_type in raw exit_statement: {:?}", raw);
            }
            None
        }

        "label_statement" => {
            let id_node = node
                .child_by_field_name("identifier")
                .or_else(|| {
                    let mut lc = node.walk();
                    node.named_children(&mut lc).next()
                })?;
            Some(Statement::Label(extract(source, id_node)))
        }

        "expression_statement" => {
            let mut ec = node.walk();
            let expr_node = node.named_children(&mut ec).next()?;
            let expr = build_expression(expr_node, source)?;
            Some(Statement::Expression(expr))
        }
        
        "on_error_statement" => {
            //println!("🎯 PARSING ON ERROR STATEMENT");
            let lower = extract(source, node).to_ascii_lowercase();
            //println!("   Node lower: {}",lower);
            for i in 0..node.child_count() {
                let child = node.child(i).unwrap();
                //println!("   Child {}: kind='{}', text='{}'", i, child.kind(), extract(source, child));
            }
            if lower.contains("resume next") {
                Some(Statement::OnError(OnErrorKind::ResumeNext))
            } else if lower.contains("goto 0") {
                Some(Statement::OnError(OnErrorKind::GoToZero))   // <-- fix
            } else if lower.contains("goto") {
                // Extract the label - it's the last identifier child
                for i in (0..node.child_count()).rev() {
                    let child = node.child(i).unwrap();
                    if child.kind() == "identifier" {
                        let label = extract(source, child);
                        //println!("   ✅ Found label: {}", label);
                        return Some(Statement::OnError(OnErrorKind::GoToLabel(label)));
                    }
                }
                //println!("   ❌ No label found, returning None");
                None
            } else {
                None
            }
        }

        "resume_statement" => {
            //println!("🎯 PARSING ON resume_statement");
            let lower = extract(source, node).to_ascii_lowercase();
            //println!("   Node lower in resume_statement: {}",lower);
            if lower.contains("resume next") {
                Some(Statement::Resume(ResumeKind::Next))
            } else if let Some(lbl) = node.child_by_field_name("label") {
                Some(Statement::Resume(ResumeKind::Label(extract(source, lbl))))
            } else {
                Some(Statement::Resume(ResumeKind::Current)) // bare Resume
            }
        }

        "enum_statement" => {
            // Extract optional visibility modifier
            let visibility = node.child_by_field_name("visibility")
                .map(|v| extract(source, v));
            
            // Extract enum name
            let name_node = node.child_by_field_name("name")?;
            let name = extract(source, name_node);
            
            eprintln!("🔨 Building Enum `{}`", name);
            
            // Extract enum members
            let mut members = Vec::new();
            let mut cursor = node.walk();
            
            for child in node.named_children(&mut cursor) {
                if child.kind() == "enum_member" {
                    if let Some(member) = build_enum_member(child, source) {
                        members.push(member);
                    }
                }
            }
            
            if members.is_empty() {
                eprintln!("⚠️ Enum `{}` has no members", name);
                return None;
            }
            
            eprintln!("✅ Built Enum `{}` with {} members", name, members.len());
            
            Some(Statement::Enum {
                visibility,
                name,
                members,
            })
        }
        "type_statement" => {
            // Extract optional visibility modifier
            let visibility = node.child_by_field_name("visibility")
                .map(|v| extract(source, v));
            
            // Extract type name
            let name_node = node.child_by_field_name("name")?;
            let name = extract(source, name_node);
            
            eprintln!("🔨 Building Type `{}`", name);
            
            // Extract type fields
            let mut fields = Vec::new();
            let mut cursor = node.walk();
            
            for child in node.named_children(&mut cursor) {
                if child.kind() == "type_member" {
                    if let Some(field) = build_type_field(child, source) {
                        fields.push(field);
                    }
                }
            }
            
            if fields.is_empty() {
                eprintln!("⚠️ Type `{}` has no fields", name);
                return None;
            }
            
            eprintln!("✅ Built Type `{}` with {} fields", name, fields.len());
            
            Some(Statement::Type {
                visibility,
                name,
                fields,
            })
        }

       "call_statement" => {
            let mut function: Option<String> = None;
            let mut args: Vec<Expression> = Vec::new();

            // only the named children: identifier, argument_list, expression
            let mut c = node.walk();
            for child in node.named_children(&mut c) {
                match child.kind() {
                    "identifier" if function.is_none() => {
                        let name = extract(source, child);
                        eprintln!("  📥 found identifier: `{}`", name);
                        function = Some(name);
                    }

                    "argument_list" => {
                        let mut ac = child.walk();
                        for expr_node in child
                            .children(&mut ac)
                            .filter(|n| n.kind() == "expression")
                        {
                            // unwrap each expression in the list
                            let mut ic = expr_node.walk();
                            let inner = expr_node.named_children(&mut ic).next().unwrap();
                            if let Some(expr) = build_expression(inner, source) {
                                eprintln!("  📥 collected arg from list: {:?}", expr);
                                args.push(expr);
                            }
                        }
                    }

                    "expression" => {
                        // unwrap the single-expr form
                        let mut ec = child.walk();
                        let inner = child
                            .named_children(&mut ec)
                            .next()
                            .expect("expression wrapper should have a child");
                        let expr = build_expression(inner, source)
                            .expect("expression child should build");
                        eprintln!("  📥 collected single-expr arg: {:?}", expr);
                        args.push(expr);
                    }

                    _ => {}
                }
            }

            let fn_name = function.unwrap_or_default();
            eprintln!("⟳ resolved function = `{}`, arg count = {}", fn_name, args.len());
            eprintln!("  ✅ emitting Call AST for `{}`\n", fn_name);

            Some(Statement::Call {
                function: fn_name,
                args,
            })
        }

        _ => {
            eprintln!("⚠️ Unhandled statement type: {} with text: {:?}", 
                     node.kind(), 
                     node.utf8_text(source.as_bytes()).unwrap_or(""));
            None
        }
    }
}

// Enhanced build_expression function to handle nested structures
fn build_expression(node: Node, source: &str) -> Option<Expression> {
    match node.kind() {
        "expression" => {
            // Unwrap the expression wrapper - look for the actual expression inside
            let mut ec = node.walk();
            if let Some(inner) = node.named_children(&mut ec).next() {
                return build_expression(inner, source);
            }
            None
        }
        
        "integer_literal" => {
            let text = extract(source, node);
            text.parse::<i64>().ok().map(Expression::Integer)
        }
        "boolean_literal" => {
            let text = extract(source, node);
            let cleaned = text.trim().to_lowercase();
            // Accept True/False (case-insensitive)
            match cleaned.as_str() {
                "true" => Some(Expression::Boolean(true)),
                "false" => Some(Expression::Boolean(false)),
                _ => None,
            }
        }
        "byte_literal" => {
            let text = extract(source, node);
            match text.parse::<u8>() {        // restrict to 0..=255
                Ok(val) => Some(Expression::Byte(val)),
                Err(_) => {
                    println!("❌ Byte literal out of range: {}", text);
                    None
                }
            }
        }
        // currency_literal
        "currency_literal" => {
            // e.g. "1234.56"
            let raw = extract(source, node);      // String lives to end of this arm
            let text = raw.trim();                // &str borrowing `raw`
            match text.parse::<f64>() {
                Ok(f) => Some(Expression::Currency(f)),
                Err(_) => {
                    eprintln!("❌ Failed to parse currency_literal: {}", text);
                    None
                }
            }
        }

        // float_literal
        "float_literal" => {
            // e.g. "3.1415926535"
            let raw = extract(source, node);      // keep the String alive
            let text = raw.trim();                // borrow it safely
            match text.parse::<f64>() {
                Ok(f) => Some(Expression::Double(f)),
                Err(_) => {
                    eprintln!("❌ Failed to parse float_literal: {}", text);
                    None
                }
            }
        }


        "date_literal" => {
            // e.g. "#10/15/2025 14:30:00#"
            let raw = extract(source, node);
            let inner = raw.trim().trim_start_matches('#').trim_end_matches('#').trim();

            // Try with time first, then date-only
            let dt = chrono::NaiveDateTime::parse_from_str(inner, "%m/%d/%Y %H:%M:%S")
                .or_else(|_| chrono::NaiveDateTime::parse_from_str(inner, "%m/%d/%Y %H:%M"))
                .ok()
                .map(|dt| dt.date());

            let d = dt.or_else(|| chrono::NaiveDate::parse_from_str(inner, "%m/%d/%Y").ok());
            match d {
                Some(date) => Some(Expression::Date(date)),
                None => {
                    eprintln!("❌ Failed to parse date_literal: {}", inner);
                    None
                }
            }
        }

        "identifier" => {
            Some(Expression::Identifier(extract(source, node)))
        }

        "parenthesized_expression" => {
            // Handle parentheses: (expr)
            //println!("  🔄 Handling parenthesized expression");
            let mut pc = node.walk();
            // Look for the expression inside parentheses
            for child in node.named_children(&mut pc) {
                if child.kind() == "expression" {
                    return build_expression(child, source);
                }
            }
            None
        }
        
        "binary_expression" => {
            let mut left_expr = None;
            let mut right_expr = None;
            let mut operator = String::new();
            
            // First, get the expressions (named children)
            let mut bc = node.walk();
            let expressions: Vec<Node> = node
                .named_children(&mut bc)
                .filter(|n| n.kind() == "expression")
                .collect();
            
            if expressions.len() >= 2 {
                left_expr = build_expression(expressions[0], source);
                right_expr = build_expression(expressions[1], source);
            }
            
            // NOW: Actually detect the operator from ALL children (including tokens)
            let mut bc2 = node.walk();
            for child in node.children(&mut bc2) {
                let child_text = child.utf8_text(source.as_bytes()).unwrap_or("");
                match child_text {
                    "+" | "-" | "*" | "/" | "&" | "<>" | ">=" | "<=" | ">" | "<" | "=" => {
                        operator = child_text.to_string();
                        break;
                    }
                    _ => continue,
                }
            }
            
            // Fallback if operator not found
            if operator.is_empty() {
                // Get the byte ranges of the two expressions
                if expressions.len() >= 2 {
                    let left_end = expressions[0].end_byte();
                    let right_start = expressions[1].start_byte();
                    
                    if right_start > left_end {
                        // Extract the text between the expressions
                        let between_text = &source[left_end..right_start];
                        eprintln!("🔍 Text between expressions: '{}'", between_text);
                        
                        // Look for operators in the between text
                        for op_char in ["+", "-", "*", "/", "&", "<>", ">=", "<=", ">", "<", "="] {
                            if between_text.contains(op_char) {
                                operator = op_char.to_string();
                                eprintln!("✅ Found operator '{}' in between text", operator);
                                break;
                            }
                        }
                    }
                }
            }
            
            // If STILL no operator found, this is definitely wrong
            if operator.is_empty() {
                eprintln!("❌ FATAL: Could not find any operator in binary expression");
                eprintln!("   This indicates a serious parsing problem");
                eprintln!("   Node: {:?}", node.utf8_text(source.as_bytes()));
                return None;  // Don't mask the problem with a default
            }
            
            if let (Some(left), Some(right)) = (left_expr, right_expr) {
                Some(Expression::BinaryOp {
                    left: Box::new(left),
                    op: operator,
                    right: Box::new(right),
                })
            } else {
                eprintln!("⚠️ Failed to build binary expression");
                None
            }
        }
        "unary_expression" => {
            let operator_node = node.child_by_field_name("operator")?;
            let operator_text = extract(source, operator_node);
            
            let argument_node = node.child_by_field_name("argument")?;
            let argument_expr = build_expression(argument_node, source)?;

            Some(Expression::UnaryOp {
                op: operator_text,
                expr: Box::new(argument_expr),
            })
        }


        
        "string_literal" => {
            let text = extract(source, node);
            // Remove opening and closing quotes
            let inner = if text.len() >= 2 {
                &text[1..text.len()-1]
            } else {
                ""
            };
            let unescaped = inner.replace("\"\"", "\"");
            Some(Expression::String(unescaped))
        }

        // ADD THIS CASE FOR PROPERTY ACCESS (ENUM MEMBER ACCESS)
        "property_access" => {
            // property_access has structure:
            //   object (identifier or expression)
            //   . (dot - not named)
            //   property (identifier)
            
            let mut obj_expr = None;
            let mut property_name = None;
            
            let mut cursor = node.walk();
            let children: Vec<Node> = node.named_children(&mut cursor).collect();
            
            // First named child is the object
            if let Some(obj_node) = children.get(0) {
                obj_expr = build_expression(*obj_node, source);
            }
            
            // Second named child is the property name
            if let Some(prop_node) = children.get(1) {
                if prop_node.kind() == "identifier" {
                    property_name = Some(extract(source, *prop_node));
                }
            }
            
            match (&obj_expr, &property_name) {
                (Some(obj), Some(prop)) => {
                    eprintln!("✅ Built PropertyAccess: obj={:?}, property={}", obj, prop);
                    Some(Expression::PropertyAccess {
                        obj: Box::new(obj.clone()),
                        property: prop.clone(),
                    })
                }
                _ => {
                    eprintln!("❌ Failed to build property_access - obj: {:?}, prop: {:?}", 
                             obj_expr, property_name);
                    None
                }
            }
        }
        
        "call_expression" => {
            // Handle function calls in expressions
            let mut function = None;
            let mut args = Vec::new();
            
            let mut cc = node.walk();
            for child in node.named_children(&mut cc) {
                match child.kind() {
                    "identifier" => {
                        if function.is_none() {
                            function = Some(extract(source, child));
                        }
                    }
                    "argument_list" => {
                        let mut ac = child.walk();
                        for arg_node in child.named_children(&mut ac) {
                            if let Some(arg_expr) = build_expression(arg_node, source) {
                                args.push(arg_expr);
                            }
                        }
                    }
                    _ => {}
                }
            }
            function.map(|f| Expression::FunctionCall { function: Box::new(Expression::Identifier(f)), args })
            // function.map(|f| Expression::Call { function: f, args })
        }
        "vba_builtin_constant" => {
            // Extract the text of the node (e.g., "vbCalGreg")
            let text = node.utf8_text(source.as_bytes()).unwrap().to_string();
            Some(Expression::BuiltInConstant(text))
        },
        "indexed_access" => {
            // Expect: identifier "(" [args] ")"
            let mut func_name: Option<String> = None;
            let mut args: Vec<Expression> = Vec::new();

            let mut wc = node.walk();
            for child in node.children(&mut wc) {
                match child.kind() {
                    "identifier" if func_name.is_none() => {
                        func_name = Some(extract(source, child));
                    }
                    "expression" => {
                        if let Some(arg) = build_expression(child, source) {
                            args.push(arg);
                        }
                    }
                    _ => {}
                }
            }

            func_name.map(|name|
                Expression::FunctionCall {
                    function: Box::new(Expression::Identifier(name)),
                    args
                }
            )
        },
        
        _ => {
            eprintln!("⚠️ Unhandled expression type: {} with text: {:?}", 
                     node.kind(), 
                     node.utf8_text(source.as_bytes()).unwrap_or(""));
            None
        }
    }
}

// Helper function for extracting text from nodes
fn extract(source: &str, node: Node) -> String {
    node.utf8_text(source.as_bytes())
        .unwrap_or("")
        .to_string()
}
impl Expression {
    pub fn from_tree_sitter_node(node: Node, source: &str) -> Result<Self, String> {
        build_expression(node, source)
            .ok_or_else(|| format!("Failed to build expression from node: {}", node.kind()))
    }
}

impl Statement {
    pub fn from_tree_sitter_node(node: Node, source: &str) -> Result<Self, String> {
        build_statement(node, source)
            .ok_or_else(|| format!("Failed to build statement from node: {}", node.kind()))
    }
}

impl ExitType {
    pub fn from_str(s: &str) -> Option<Self> {
        match s.to_ascii_lowercase().as_str() {
            "sub"      => Some(ExitType::Sub),
            "for"      => Some(ExitType::For),
            "do"       => Some(ExitType::Do),
            "while"    => Some(ExitType::While),
            "function" => Some(ExitType::Function),
            "property" => Some(ExitType::Property),
            "select"   => Some(ExitType::Select),
            _ => None,
        }
    }
    
    pub fn as_str(&self) -> &'static str {
        match self {
            ExitType::For      => "For",
            ExitType::Do       => "Do",
            ExitType::While    => "While",
            ExitType::Sub      => "Sub",
            ExitType::Function => "Function",
            ExitType::Property => "Property",
            ExitType::Select   => "Select",
        }
    }
} // End of impl ExitType

// Add new helper function to build enum members:
fn build_enum_member(node: Node, source: &str) -> Option<EnumMember> {
    // Extract member name
    let name_node = node.child_by_field_name("name")?;
    let name = extract(source, name_node);
    
    // Extract optional value
    let value = node.child_by_field_name("value")
        .and_then(|v| build_expression(v, source));
    
    Some(EnumMember { name, value })
}
// Add new helper function to build type fields:
fn build_type_field(node: Node, source: &str) -> Option<TypeField> {
    // Extract field name
    let name_node = node.child_by_field_name("name")?;
    let name = extract(source, name_node);
    
    // Extract field type
    let type_node = node.child_by_field_name("type")?;
    let field_type = extract(source, type_node);
    
    // Extract optional array dimensions
    let dimensions = node.child_by_field_name("dimensions")
        .map(|dims_node| build_array_dimensions(dims_node, source));
    
    // Extract optional string length (for String * length)
    let string_length = node.child_by_field_name("string_length")
        .and_then(|len_node| {
            let len_text = extract(source, len_node);
            len_text.parse::<i64>().ok()
        });
    
    eprintln!("  ✅ Built field: {} As {}{}", 
             name, 
             field_type,
             if let Some(len) = string_length { format!(" * {}", len) } else { String::new() }
    );
    
    Some(TypeField {
        name,
        field_type,
        dimensions,
        string_length,
    })
}

// Add helper to build array dimensions:
fn build_array_dimensions(node: Node, source: &str) -> Vec<ArrayDimension> {
    let mut dimensions = Vec::new();
    let mut cursor = node.walk();
    
    for child in node.named_children(&mut cursor) {
        if child.kind() == "array_dimension" {
            if let Some(dim) = build_array_dimension(child, source) {
                dimensions.push(dim);
            }
        }
    }
    
    dimensions
}

fn build_array_dimension(node: Node, source: &str) -> Option<ArrayDimension> {
    let lower_node = node.child_by_field_name("lower");
    let upper_node = node.child_by_field_name("upper")
        .or_else(|| node.child_by_field_name("size"))?;
    
    let lower = lower_node.and_then(|n| build_expression(n, source));
    let upper = build_expression(upper_node, source)?;
    
    Some(ArrayDimension { lower, upper })
}




// "exit_statement" => {
//     // Preferred path: use the grammar field if present.
//     if let Some(exit_type_node) = node.child_by_field_name("exit_type") {
//         let exit_type_str = extract(source, exit_type_node).trim().to_string();
//         if let Some(exit_type) = ExitType::from_str(&exit_type_str) {
//             return Some(Statement::Exit(exit_type));
//         } else {
//             eprintln!("Unknown exit type (field): {}", exit_type_str);
//             // fall through to raw-text fallback
//         }
//     }

//     // Fallback: parse raw node text like "Exit For", "Exit Do", … even if the field is missing.
//     let raw = extract(source, node);
//     // Grab the token after "Exit" (handles extra whitespace and a trailing ':' if present)
//     if let Some(tok) = raw.split_whitespace().nth(1) {
//         let cleaned = tok.trim_end_matches(':').trim();
//         if let Some(exit_type) = ExitType::from_str(cleaned) {
//             return Some(Statement::Exit(exit_type));
//         } else {
//             eprintln!("Unknown exit type (raw): {}", cleaned);
//         }
//     } else {
//         eprintln!("Missing exit_type in raw exit_statement: {:?}", raw);
//     }
//     None
// }



// "exit_statement" => {
//     if let Some(exit_type_node) = node.child_by_field_name("exit_type") {
//         let exit_type_str = extract(source, exit_type_node).trim().to_string();
//         if let Some(exit_type) = ExitType::from_str(&exit_type_str) {
//             println!("in build exit statement");
//             Some(Statement::Exit(exit_type))
//         } else {
//             eprintln!("Unknown exit type: {}", exit_type_str);
//             None
//         }
//     } else {
//         eprintln!("Missing exit_type node in exit_statement");
//         None
//     }
// }