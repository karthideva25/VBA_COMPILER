use crate::ast::{Statement, ForStatement, DoWhileStatement, Expression, ExitType, OnErrorKind, ResumeKind, EnumMember,TypeField, DoWhileConditionType};
use crate::interpreter::evaluate_expression;
use crate::context::{Context, Value, ScopeKind, FieldDefinition, ErrObject, OnErrorMode,DeclaredType};
use crate::interpreter::builtins::handle_builtin_call_bool;
use crate::interpreter::coerce::coerce_to_declared;
use std::collections::HashMap;

// === Control flow signals used internally by the interpreter ===
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ControlFlow {
    Continue,
    ExitSub,
    ExitFunction,
    ExitProperty,
    ExitFor,
    ContinueFor,
    ExitDo,
    ContinueDo,
    ExitWhile,
    ContinueWhile,
    ExitSelect,
    GoToLabel(String),
    ResumeNext,      // On Error Resume Next, or Resume Next
    ResumeCurrent,   // bare Resume (re-exec the failing line)
}

impl ControlFlow {
    pub fn from_exit_type(et: &crate::ast::ExitType) -> Self {
        use crate::ast::ExitType::*;
        match et {
            Sub      => ControlFlow::ExitSub,
            Function => ControlFlow::ExitFunction,
            Property => ControlFlow::ExitProperty,
            For      => ControlFlow::ExitFor,
            Do       => ControlFlow::ExitDo,
            While    => ControlFlow::ExitWhile,
            Select   => ControlFlow::ExitSelect,
        }
    }
}

// ————————————————————————————————————————————————————————————————————————
// Execute a single statement, returning a control-flow signal.
// IMPORTANT: `pc` is the index of this statement inside the current list.
// ————————————————————————————————————————————————————————————————————————
pub(crate) fn execute_statement(stmt: &Statement, ctx: &mut Context, pc: usize) -> ControlFlow {
    //println!("🔍 execute_statement called with: {:?}", stmt);
    match stmt {
        Statement::BlankLine => ControlFlow::Continue,
        

        Statement::Comment(text) => {
            //println!("Comment: {}", text);
            ControlFlow::Continue
        }

        Statement::OptionExplicit => {
            ctx.enable_option_explicit();
            ControlFlow::Continue
        }

        // Record subroutines for later calls
        Statement::Subroutine { name, params, body } => {
            ctx.define_sub(name.clone(), params.clone(), body.clone());
            ctx.log(&format!("Defined subroutine {}", name));
            ControlFlow::Continue
        }

        Statement::Dim { names } => {
            for (v, maybe_type) in names {
                // Register this variable as declared (for Option Explicit)
                ctx.declare_variable(v);
                
                let initial_value = if let Some(type_name) = maybe_type {
                    // First check if it's a user-defined type
                    if ctx.is_type_defined(type_name) {
                        match ctx.create_type_instance(type_name) {
                            Some(instance) => {
                                // Set the type in context
                                ctx.set_var_type(v.clone(), crate::context::DeclaredType::Variant);
                                instance
                            }
                            None => {
                                // fallback to empty string
                                ctx.set_var_type(v.clone(), crate::context::DeclaredType::Variant);
                                Value::String(String::new())
                            }
                        }
                    } else {
                        let ty = crate::context::DeclaredType::from_opt_str(Some(type_name));
                        ctx.set_var_type(v.clone(), ty);
                        match ty {
                            crate::context::DeclaredType::Byte     => Value::Byte(0),
                            crate::context::DeclaredType::Integer  => Value::Integer(0),
                            crate::context::DeclaredType::Long     => Value::Long(0),
                            crate::context::DeclaredType::LongLong => Value::LongLong(0),
                            crate::context::DeclaredType::Object   => Value::Object(None), 
                            crate::context::DeclaredType::Currency => Value::Currency(0.0),
                            crate::context::DeclaredType::Date     => chrono::NaiveDate::from_ymd_opt(1899,12,30).map(Value::Date).unwrap_or(Value::Date(chrono::NaiveDate::from_ymd_opt(1899,12,30).unwrap())),
                            crate::context::DeclaredType::Double   => Value::Double(0.0),
                            crate::context::DeclaredType::Decimal  => Value::Decimal(0.0),
                            crate::context::DeclaredType::Single   => Value::Single(0.0),
                            crate::context::DeclaredType::String   => Value::String(String::new()),
                            crate::context::DeclaredType::Boolean  => Value::Boolean(false),
                            crate::context::DeclaredType::Variant  => Value::String(String::new()),
                        }
                    }
                } else {
                    // No type specified - default to Variant
                    ctx.set_var_type(v.clone(), crate::context::DeclaredType::Variant);
                    Value::String(String::new())
                };
                ctx.declare_local(v.clone(), initial_value);
            }
            ControlFlow::Continue
        }
        
        Statement::ReDim { preserve, arrays } => {
            for redim_array in arrays {
                // Calculate total size from dimensions
                let mut total_size = 1usize;
                let mut dim_sizes = Vec::new();
                
                for dim in &redim_array.dimensions {
                    let lower = if let Some(ref lower_expr) = dim.lower {
                        match eval_opt(lower_expr, ctx) {
                            Some(v) => match value_to_integer(&v) {
                                Ok(n) => n,
                                Err(_) => {
                                    ctx.log(&format!("Error: Invalid lower bound in ReDim"));
                                    return raise_runtime_error(ctx, 13, "Type mismatch in ReDim lower bound", pc);
                                }
                            },
                            None => 0,
                        }
                    } else {
                        0
                    };
                    
                    let upper = match eval_opt(&dim.upper, ctx) {
                        Some(v) => match value_to_integer(&v) {
                            Ok(n) => n,
                            Err(_) => {
                                ctx.log(&format!("Error: Invalid upper bound in ReDim"));
                                return raise_runtime_error(ctx, 13, "Type mismatch in ReDim upper bound", pc);
                            }
                        },
                        None => {
                            ctx.log(&format!("Error: Could not evaluate upper bound in ReDim"));
                            return raise_runtime_error(ctx, 13, "Invalid ReDim upper bound", pc);
                        }
                    };
                    
                    let size = (upper - lower + 1) as usize;
                    dim_sizes.push((lower, upper, size));
                    total_size *= size;
                }
                
                // Get existing array if preserve is true
                let existing_data = if *preserve {
                    ctx.get_var(&redim_array.name)
                } else {
                    None
                };
                
                // Create new array
                let new_array = if *preserve {
                    // Preserve existing data
                    if let Some(Value::Array(old_arr)) = existing_data {
                        let mut new_arr = vec![Value::Integer(0); total_size];
                        
                        // Copy old data (only what fits in new dimensions)
                        let copy_len = old_arr.len().min(total_size);
                        for i in 0..copy_len {
                            new_arr[i] = old_arr[i].clone();
                        }
                        
                        Value::Array(new_arr)
                    } else {
                        // Variable wasn't an array before, create new
                        Value::Array(vec![Value::Integer(0); total_size])
                    }
                } else {
                    // Don't preserve, create fresh array
                    Value::Array(vec![Value::Integer(0); total_size])
                };
                
                ctx.set_var(redim_array.name.clone(), new_array);
                
                let preserve_str = if *preserve { " Preserve" } else { "" };
                ctx.log(&format!("ReDim{} {} with size {}", 
                                preserve_str, 
                                redim_array.name, 
                                total_size));
            }
            
            ControlFlow::Continue
        }

        // SET/Assignment
        Statement::Set { target, expr } => {
            if let Some(val) = eval_opt(expr, ctx) {
                ctx.set_var(target.clone(), val);
            }
            ControlFlow::Continue
        }
        
       Statement::Assignment { lvalue, rvalue } => {
            // 1) Evaluate the RHS expression safely, catching interpreter errors
            let rhs_val_res = crate::interpreter::evaluate_expression(rvalue, ctx);

            if let Err(e) = rhs_val_res.as_ref() {
                // Capture the runtime error into the VBA Err object
                ctx.err = Some(ErrObject {
                    number: 13,
                    description: e.to_string(),
                    source: "Interpreter".into(),
                });
            }

            // 2) If any error is active, handle it according to On Error settings
            if ctx.err.is_some() {
                if let Some(flow) = maybe_handle_error(ctx, pc) {
                    return flow;
                }
                return ControlFlow::Continue;
            }

            // 3) Safe unwrap – expression evaluated successfully
            let rhs_val = match rhs_val_res {
                Ok(v) => v,
                Err(_) => return ControlFlow::Continue,
            };

            // 4) Now perform the actual assignment
            match lvalue {
                crate::ast::AssignmentTarget::PropertyAccess { object, property } => {
                    // Check if object variable is declared (Option Explicit)
                    if let Err(e) = ctx.validate_variable_usage(object) {
                        ctx.log(&e);
                        ctx.err = Some(ErrObject {
                            number: 451, // VBA error: Variable not defined
                            description: e,
                            source: "Interpreter".into(),
                        });
                        if let Some(flow) = maybe_handle_error(ctx, pc) {
                            return flow;
                        }
                        return ControlFlow::Continue;
                    }
                    
                    if let Some(mut obj_val) = ctx.get_var(object) {
                        match obj_val.set_field(property, rhs_val.clone()) {
                            Ok(()) => {
                                ctx.set_var(object.to_string(), obj_val);
                                ctx.log(&format!("Set {}.{} = {}", object, property, rhs_val.as_string()));
                            }
                            Err(e) => {
                                ctx.log(&format!("Error setting field: {}", e));
                                ctx.err = Some(ErrObject {
                                    number: 13,
                                    description: format!("Error setting field: {}", e),
                                    source: "Interpreter".into(),
                                });
                                if let Some(flow) = maybe_handle_error(ctx, pc) {
                                    return flow;
                                }
                                return ControlFlow::Continue;
                            }
                        }
                    } else {
                        ctx.log(&format!("Error: Variable '{}' not found", object));
                        ctx.err = Some(ErrObject {
                            number: 91,
                            description: format!("Variable '{}' not found", object),
                            source: "Interpreter".into(),
                        });
                        if let Some(flow) = maybe_handle_error(ctx, pc) {
                            return flow;
                        }
                        return ControlFlow::Continue;
                    }
                }

                crate::ast::AssignmentTarget::Identifier(var_name) => {
                    // Check if variable is declared when Option Explicit is enabled
                    if let Err(e) = ctx.validate_variable_usage(var_name) {
                        ctx.log(&e);
                        ctx.err = Some(ErrObject {
                            number: 451, // VBA error: Variable not defined
                            description: e,
                            source: "Interpreter".into(),
                        });
                        if let Some(flow) = maybe_handle_error(ctx, pc) {
                            return flow;
                        }
                        return ControlFlow::Continue;
                    }
                    
                    if let Some(ty) = ctx.get_var_type(var_name) {
                        match crate::interpreter::coerce::coerce_to_declared(rhs_val, ty) {
                            Ok(v) => {
                                ctx.set_var(var_name.clone(), v);
                            }
                            Err(e) => {
                                ctx.log(&format!("Type mismatch assigning to {}: {}", var_name, e));
                                ctx.err = Some(ErrObject {
                                    number: 13,
                                    description: format!("Type mismatch assigning to {}: {}", var_name, e),
                                    source: "Interpreter".into(),
                                });
                                if let Some(flow) = maybe_handle_error(ctx, pc) {
                                    return flow;
                                }
                                return ControlFlow::Continue;
                            }
                        }
                    } else {
                        // No declared type => Variant semantics
                        ctx.set_var(var_name.clone(), rhs_val);
                    }
                }

                // ✅ ADD THIS NEW CASE FOR INDEXED ACCESS (array assignment)
                crate::ast::AssignmentTarget::IndexedAccess { array, indices } => {
                    // Check if array variable is declared (Option Explicit)
                    if let Err(e) = ctx.validate_variable_usage(array) {
                        ctx.log(&e);
                        ctx.err = Some(ErrObject {
                            number: 451, // VBA error: Variable not defined
                            description: e,
                            source: "Interpreter".into(),
                        });
                        if let Some(flow) = maybe_handle_error(ctx, pc) {
                            return flow;
                        }
                        return ControlFlow::Continue;
                    }

                    // Get the array from context
                    let array_value = match ctx.get_var(array) {
                        Some(val) => val,
                        None => {
                            ctx.log(&format!("Error: Array '{}' not found", array));
                            ctx.err = Some(ErrObject {
                                number: 9, // VBA error: Subscript out of range
                                description: format!("Array '{}' not found", array),
                                source: "Interpreter".into(),
                            });
                            if let Some(flow) = maybe_handle_error(ctx, pc) {
                                return flow;
                            }
                            return ControlFlow::Continue;
                        }
                    };

                    // Evaluate all indices
                    let mut index_values = Vec::new();
                    for idx_expr in indices {
                        match crate::interpreter::evaluate_expression(idx_expr, ctx) {
                            Ok(idx_val) => {
                                match value_to_integer(&idx_val) {
                                    Ok(idx) => index_values.push(idx as usize),
                                    Err(_) => {
                                        ctx.log(&format!("Error: Array index must be numeric"));
                                        ctx.err = Some(ErrObject {
                                            number: 13, // Type mismatch
                                            description: "Array index must be numeric".to_string(),
                                            source: "Interpreter".into(),
                                        });
                                        if let Some(flow) = maybe_handle_error(ctx, pc) {
                                            return flow;
                                        }
                                        return ControlFlow::Continue;
                                    }
                                }
                            }
                            Err(e) => {
                                ctx.log(&format!("Error evaluating array index: {}", e));
                                ctx.err = Some(ErrObject {
                                    number: 13,
                                    description: format!("Error evaluating array index: {}", e),
                                    source: "Interpreter".into(),
                                });
                                if let Some(flow) = maybe_handle_error(ctx, pc) {
                                    return flow;
                                }
                                return ControlFlow::Continue;
                            }
                        }
                    }

                    // Handle array assignment (for now, only 1D arrays)
                    if index_values.len() != 1 {
                        ctx.log(&format!("Error: Multi-dimensional arrays not yet fully supported"));
                        ctx.err = Some(ErrObject {
                            number: 9,
                            description: "Multi-dimensional arrays not yet fully supported".to_string(),
                            source: "Interpreter".into(),
                        });
                        if let Some(flow) = maybe_handle_error(ctx, pc) {
                            return flow;
                        }
                        return ControlFlow::Continue;
                    }

                    let index = index_values[0];
                    match array_value {
                        Value::Array(mut arr) => {
                            if index >= arr.len() {
                                ctx.log(&format!("Error: Array index {} out of bounds (size: {})", index, arr.len()));
                                ctx.err = Some(ErrObject {
                                    number: 9, // Subscript out of range
                                    description: format!("Array index {} out of bounds", index),
                                    source: "Interpreter".into(),
                                });
                                if let Some(flow) = maybe_handle_error(ctx, pc) {
                                    return flow;
                                }
                                return ControlFlow::Continue;
                            }
                            
                            // Set the array element
                            arr[index] = rhs_val.clone();
                            ctx.set_var(array.clone(), Value::Array(arr));
                            ctx.log(&format!("Set {}({}) = {}", array, index, rhs_val.as_string()));
                        }
                        _ => {
                            ctx.log(&format!("Error: '{}' is not an array", array));
                            ctx.err = Some(ErrObject {
                                number: 13, // Type mismatch
                                description: format!("'{}' is not an array", array),
                                source: "Interpreter".into(),
                            });
                            if let Some(flow) = maybe_handle_error(ctx, pc) {
                                return flow;
                            }
                            return ControlFlow::Continue;
                        }
                    }
                }
            }

            ControlFlow::Continue
        }
        

        Statement::MsgBox { expr } => {
            if let Some(val) = eval_opt(expr, ctx) {
                ctx.log(&to_string(&val));
            }
            ControlFlow::Continue
        }

        Statement::GoTo { label } => ControlFlow::GoToLabel(label.clone()),

        // If/ElseIf/Else: delegate to nested statement lists so they get their own PC
        Statement::If { condition, then_branch, else_if, else_branch } => {
            if let Some(cv) = eval_opt(condition, ctx) {
                if is_truthy(&cv) {
                    return execute_statement_list(then_branch, ctx);
                }
                for (elseif_cond, elseif_stmts) in else_if {
                    if let Some(elseif_val) = eval_opt(elseif_cond, ctx) {
                        if is_truthy(&elseif_val) {
                            return execute_statement_list(elseif_stmts, ctx);
                        }
                    }
                }
                if !else_branch.is_empty() {
                    return execute_statement_list(else_branch, ctx);
                }
            }
            ControlFlow::Continue
        }

        Statement::For(for_stmt) => execute_for_loop(for_stmt, ctx, pc),
        Statement::DoWhile(do_stmt) => execute_do_while_loop(do_stmt, ctx, pc),

        Statement::Exit(exit_type) => ControlFlow::from_exit_type(exit_type),

        Statement::Label(_) => ControlFlow::Continue,

        Statement::Expression(expr) => {
            let _ = eval_opt(expr, ctx);
            ControlFlow::Continue
        }

        // ——— Error handling directives
        Statement::OnError(kind) => {
            //println!("🎯🎯🎯 OnError statement detected!");
            match kind {
                OnErrorKind::ResumeNext     => { 
                    //println!("   Setting mode to ResumeNext");
                    ctx.on_error_mode = OnErrorMode::ResumeNext; ctx.on_error_label = None; }
                OnErrorKind::GoToLabel(lbl) => { 
                    //println!("   Setting mode to GoTo, label: {}", lbl);
                    ctx.on_error_mode = OnErrorMode::GoTo;       ctx.on_error_label = Some(lbl.clone()); }
                OnErrorKind::GoToZero       => {  //
                    //println!("   Setting mode to None");
                    ctx.on_error_mode = OnErrorMode::None;       ctx.on_error_label = None; }
            }
            //println!("   New mode: {:?}, label: {:?}", ctx.on_error_mode, ctx.on_error_label);
            ControlFlow::Continue
        }

        // ——— Resume statements
        Statement::Resume(kind) => {
            match kind {
                ResumeKind::Current    => ControlFlow::ResumeCurrent,
                ResumeKind::Next       => ControlFlow::ResumeNext,
                ResumeKind::Label(lbl) => ControlFlow::GoToLabel(lbl.clone()),
            }
        }

        // ——— Subroutine call with scoping
        // Statement::Call { function, args } => {
        //     if handle_builtin_call_bool(function, args, ctx) {
        //         return ControlFlow::Continue;
        //     }

        //     let (params, body) = match ctx.subs.get(function).cloned() {
        //         Some(pb) => pb,
        //         None => {
        //             ctx.log(&format!("*** Call `{}` not implemented", function));
        //             return ControlFlow::Continue;
        //         }
        //     };

        //     if params.len() != args.len() {
        //         ctx.log(&format!(
        //             "*** Error: Sub `{}` expects {} arguments, got {}",
        //             function, params.len(), args.len()
        //         ));
        //         return ControlFlow::Continue;
        //     }

        //     let mut arg_vals = Vec::with_capacity(args.len());
        //     for (idx, a) in args.iter().enumerate() {
        //         match eval_opt(a, ctx) {
        //             Some(v) => arg_vals.push(v),
        //             None => {
        //                 ctx.log(&format!("*** Error: could not evaluate argument {} for `{}`", idx + 1, function));
        //                 return ControlFlow::Continue;
        //             }
        //         }
        //     }

        //     ctx.log(&format!("Entering Sub {}", function));
        //     ctx.push_scope(function.clone(), ScopeKind::Subroutine);
        //     for (param, val) in params.iter().zip(arg_vals.into_iter()) {
        //         ctx.declare_local(param.clone(), val);
        //     }

        //     let flow = execute_statement_list(&body, ctx);

        //     ctx.pop_scope();
        //     ctx.log(&format!("Leaving Sub {}", function));

        //     match flow {
        //         ControlFlow::Continue
        //         | ControlFlow::ExitSub
        //         | ControlFlow::ExitFunction
        //         | ControlFlow::ExitProperty
        //         | ControlFlow::ExitFor
        //         | ControlFlow::ContinueFor
        //         | ControlFlow::ExitDo
        //         | ControlFlow::ContinueDo
        //         | ControlFlow::ExitWhile
        //         | ControlFlow::ContinueWhile
        //         | ControlFlow::ExitSelect
        //         | ControlFlow::GoToLabel(_)
        //         | ControlFlow::ResumeNext
        //         | ControlFlow::ResumeCurrent => ControlFlow::Continue,
        //     }
        // }

        // In the Call statement handler, update parameter declaration:
        Statement::Call { function, args } => {
            if handle_builtin_call_bool(function, args, ctx) {
                return ControlFlow::Continue;
            }

            let (params, body) = match ctx.subs.get(function).cloned() {
                Some(pb) => pb,
                None => {
                    ctx.log(&format!("*** Call `{}` not implemented", function));
                    return ControlFlow::Continue;
                }
            };

            if params.len() != args.len() {
                ctx.log(&format!(
                    "*** Error: Sub `{}` expects {} arguments, got {}",
                    function, params.len(), args.len()
                ));
                return ControlFlow::Continue;
            }

            let mut arg_vals = Vec::with_capacity(args.len());
            for (idx, a) in args.iter().enumerate() {
                match eval_opt(a, ctx) {
                    Some(v) => arg_vals.push(v),
                    None => {
                        ctx.log(&format!("*** Error: could not evaluate argument {} for `{}`", idx + 1, function));
                        return ControlFlow::Continue;
                    }
                }
            }

            ctx.log(&format!("Entering Sub {}", function));
            ctx.push_scope(function.clone(), ScopeKind::Subroutine);
            
            // Declare parameters in the new scope (important for Option Explicit)
            for (param, val) in params.iter().zip(arg_vals.into_iter()) {
                ctx.declare_variable(param);  // ADD THIS
                ctx.declare_local(param.clone(), val);
            }

            let flow = execute_statement_list(&body, ctx);

            ctx.pop_scope();
            ctx.log(&format!("Leaving Sub {}", function));

            match flow {
                ControlFlow::Continue
                | ControlFlow::ExitSub
                | ControlFlow::ExitFunction
                | ControlFlow::ExitProperty
                | ControlFlow::ExitFor
                | ControlFlow::ContinueFor
                | ControlFlow::ExitDo
                | ControlFlow::ContinueDo
                | ControlFlow::ExitWhile
                | ControlFlow::ContinueWhile
                | ControlFlow::ExitSelect
                | ControlFlow::GoToLabel(_)
                | ControlFlow::ResumeNext
                | ControlFlow::ResumeCurrent => ControlFlow::Continue,
            }
        }

        Statement::Enum { visibility, name, members } => {
            execute_enum_statement(visibility.as_deref(), name, members, ctx)
        }
        Statement::Type { visibility, name, fields } => {
            execute_type_statement(visibility.as_deref(), name, fields, ctx)
        }
    }
}

/// Execute a list of statements until completion or control transfer.
pub fn execute_statement_list(stmts: &[Statement], ctx: &mut Context) -> ControlFlow {
    // Pre-index labels
    let mut labels = HashMap::<String, usize>::new();
    for (idx, s) in stmts.iter().enumerate() {
        if let Statement::Label(name) = s {
            labels.insert(name.clone(), idx);
        }
    }

    let mut i = 0usize;
    while i < stmts.len() {
        //println!("\n▶️  Executing statement {} of {}", i, stmts.len());
        
        let flow = execute_statement(&stmts[i], ctx, i);
        //println!("◀️  Statement returned: {:?}", flow);
        match flow {
            ControlFlow::Continue => i += 1,

            ControlFlow::ResumeCurrent => {
                //println!("   🔄 Processing ResumeCurrent");
                if !ctx.resume_valid {
                    return raise_runtime_error(ctx, 20, "Invalid Resume", i);
                }
                if let Some(pc) = ctx.resume_pc {
                    ctx.resume_valid = false;
                    i = pc; // re-exec faulting statement
                } else {
                    return raise_runtime_error(ctx, 20, "Resume without error", i);
                }
            }

            ControlFlow::ResumeNext => {
                if !ctx.resume_valid {
                    return raise_runtime_error(ctx, 20, "Invalid Resume Next", i);
                }
                if let Some(pc) = ctx.resume_pc {
                    ctx.resume_valid = false;
                    //println!("   🔄 Continuing at statement {}", pc + 1);
                    i = pc + 1; // continue after faulting statement
                } else {
                    return raise_runtime_error(ctx, 20, "Resume Next without error", i);
                }
            }

            ControlFlow::GoToLabel(lbl) => {
                //println!("   🎯 Processing GoTo: {}", lbl);
                
                if let Some(&dest) = labels.get(&lbl) {
                    // jumping invalidates armed resume
                    ctx.resume_valid = false;
                    // println!("   🎯 Jumping to statement {}", dest);
                    i = dest + 1;
                } else {
                    // println!("   🎯 Label not in this scope, bubbling up");
                    // return raise_runtime_error(ctx, 35, "Label not defined", i);
                    return ControlFlow::GoToLabel(lbl);
                }
            }

            // other => {
                // println!("   ⬆️  Bubbling up control flow: {:?}", other);
            ControlFlow::ExitSub
            | ControlFlow::ExitFunction
            | ControlFlow::ExitProperty
            | ControlFlow::ExitFor
            | ControlFlow::ContinueFor
            | ControlFlow::ExitDo
            | ControlFlow::ContinueDo
            | ControlFlow::ExitWhile
            | ControlFlow::ContinueWhile
            | ControlFlow::ExitSelect => {
                    return flow;
            }
                // return other;
            // } // Exit*/Continue* bubble outward
        }
    }
    ControlFlow::Continue
}

/// Minimal `For` loop driver.
fn execute_for_loop(for_stmt: &ForStatement, ctx: &mut Context, pc: usize) -> ControlFlow {
    // Evaluate bounds
    let start_val = match eval_opt(&for_stmt.start, ctx) {
        Some(v) => v,
        None => return raise_runtime_error(ctx, 13, "Type mismatch in For start", pc),
    };

    let end_val = match eval_opt(&for_stmt.end, ctx) {
        Some(v) => v,
        None => return raise_runtime_error(ctx, 13, "Type mismatch in For end", pc),
    };

    let step_val = if let Some(step_expr) = &for_stmt.step {
        match eval_opt(step_expr, ctx) {
            Some(v) => v,
            None => return raise_runtime_error(ctx, 13, "Type mismatch in For Step", pc),
        }
    } else {
        Value::Integer(1)
    };

    // Coerce
    let start_int = match value_to_integer(&start_val) {
        Ok(n) => n,
        Err(_) => return raise_runtime_error(ctx, 13, "Type mismatch in For start", pc),
    };
    let end_int = match value_to_integer(&end_val) {
        Ok(n) => n,
        Err(_) => return raise_runtime_error(ctx, 13, "Type mismatch in For end", pc),
    };
    let step_int = match value_to_integer(&step_val) {
        Ok(n) => n,
        Err(_) => return raise_runtime_error(ctx, 13, "Type mismatch in For Step", pc),
    };

    if step_int == 0 {
        return raise_runtime_error(ctx, 6, "For Step cannot be zero", pc);
    }

    // Initialize loop counter
    let mut counter = start_int;
    ctx.set_var(for_stmt.counter.clone(), Value::Integer(counter));
    //println!("\n🔁 === FOR LOOP START: {} from {} to {} step {} ===", 
            //for_stmt.counter, start_int, end_int, step_int);
    loop {
        let cond_ok = if step_int > 0 { counter <= end_int } else { counter >= end_int };
        if !cond_ok {
            //println!("🔁 FOR LOOP END: condition false (counter={})", counter);
            break;
        }
        //println!("\n🔁 --- For iteration: {} = {} ---", for_stmt.counter, counter);
       
        match execute_statement_list(&for_stmt.body, ctx) {
            ControlFlow::Continue => { 
                //println!("🔁 Loop body returned Continue");
            /* keep looping */ }

            ControlFlow::ExitFor      => {
                //println!("🔁 ExitFor encountered");
                return ControlFlow::Continue;
            }
            ControlFlow::ContinueFor  => {  /* step and re-check */ }

            ControlFlow::ResumeNext   => { /* already advanced by list */ }
            ControlFlow::GoToLabel(s) =>{
                //println!("🔁 GoToLabel encountered: {}", s);
                 return ControlFlow::GoToLabel(s);}

            ControlFlow::ExitDo        => return ControlFlow::ExitDo,
            ControlFlow::ContinueDo    => return ControlFlow::ContinueDo,
            ControlFlow::ExitWhile     => return ControlFlow::ExitWhile,
            ControlFlow::ContinueWhile => return ControlFlow::ContinueWhile,
            ControlFlow::ExitSelect    => return ControlFlow::ExitSelect,

            ControlFlow::ExitSub       => return ControlFlow::ExitSub,
            ControlFlow::ExitFunction  => return ControlFlow::ExitFunction,
            ControlFlow::ExitProperty  => return ControlFlow::ExitProperty,

            ControlFlow::ResumeCurrent => return ControlFlow::ResumeCurrent,
        }

        // Step
        counter += step_int;
        //println!("🔁 Stepping: {} = {}", for_stmt.counter, counter);
        ctx.set_var(for_stmt.counter.clone(), Value::Integer(counter));
    }

    ControlFlow::Continue
}

fn execute_do_while_loop(do_stmt: &DoWhileStatement, ctx: &mut Context, pc: usize) -> ControlFlow {
    use DoWhileConditionType::*;
    
    eprintln!("\n🔁 === DO WHILE LOOP START: type={:?}, test_at_end={} ===", 
             do_stmt.condition_type, do_stmt.test_at_end);
    
    // Helper to evaluate condition
    let should_continue = |ctx: &mut Context| -> Result<bool, ControlFlow> {
        match &do_stmt.condition {
            Some(cond_expr) => {
                match eval_opt(cond_expr, ctx) {
                    Some(val) => {
                        let truthy = is_truthy(&val);
                        match do_stmt.condition_type {
                            While => Ok(truthy),           // Continue while true
                            Until => Ok(!truthy),          // Continue until true (i.e., while false)
                            DoWhileConditionType::Infinite => Ok(true),  // Should not happen, but handle it
                        }
                    }
                    Option::None => {  // Explicitly use Option::None
                        Err(raise_runtime_error(ctx, 13, "Type mismatch in Do loop condition", pc))
                    }
                }
            }
            Option::None => Ok(true), // Infinite loop - explicitly use Option::None
        }
    };
    
    // Pre-test loops: Do While/Until...Loop
    if !do_stmt.test_at_end {
        loop {
            // Check condition at start
            match should_continue(ctx) {
                Ok(true) => { /* continue */ }
                Ok(false) => {
                    eprintln!("🔁 DO LOOP END: pre-condition false");
                    break;
                }
                Err(flow) => return flow,
            }
            
            eprintln!("🔁 --- Do loop iteration (pre-test) ---");
            
            match execute_statement_list(&do_stmt.body, ctx) {
                ControlFlow::Continue => { /* keep looping */ }
                
                ControlFlow::ExitDo => {
                    eprintln!("🔁 ExitDo encountered");
                    return ControlFlow::Continue;
                }
                
                ControlFlow::ContinueDo => { /* continue to next iteration */ }
                
                ControlFlow::GoToLabel(s) => {
                    eprintln!("🔁 GoToLabel encountered: {}", s);
                    return ControlFlow::GoToLabel(s);
                }
                
                ControlFlow::ResumeNext => { /* already advanced by list */ }
                
                // Bubble up other control flows
                ControlFlow::ExitFor       => return ControlFlow::ExitFor,
                ControlFlow::ContinueFor   => return ControlFlow::ContinueFor,
                ControlFlow::ExitWhile     => return ControlFlow::ExitWhile,
                ControlFlow::ContinueWhile => return ControlFlow::ContinueWhile,
                ControlFlow::ExitSelect    => return ControlFlow::ExitSelect,
                ControlFlow::ExitSub       => return ControlFlow::ExitSub,
                ControlFlow::ExitFunction  => return ControlFlow::ExitFunction,
                ControlFlow::ExitProperty  => return ControlFlow::ExitProperty,
                ControlFlow::ResumeCurrent => return ControlFlow::ResumeCurrent,
            }
        }
    } 
    // Post-test loops: Do...Loop While/Until
    else {
        loop {
            eprintln!("🔁 --- Do loop iteration (post-test) ---");
            
            match execute_statement_list(&do_stmt.body, ctx) {
                ControlFlow::Continue => { /* keep looping */ }
                
                ControlFlow::ExitDo => {
                    eprintln!("🔁 ExitDo encountered");
                    return ControlFlow::Continue;
                }
                
                ControlFlow::ContinueDo => { /* check condition and continue */ }
                
                ControlFlow::GoToLabel(s) => {
                    eprintln!("🔁 GoToLabel encountered: {}", s);
                    return ControlFlow::GoToLabel(s);
                }
                
                ControlFlow::ResumeNext => { /* already advanced by list */ }
                
                // Bubble up other control flows
                ControlFlow::ExitFor       => return ControlFlow::ExitFor,
                ControlFlow::ContinueFor   => return ControlFlow::ContinueFor,
                ControlFlow::ExitWhile     => return ControlFlow::ExitWhile,
                ControlFlow::ContinueWhile => return ControlFlow::ContinueWhile,
                ControlFlow::ExitSelect    => return ControlFlow::ExitSelect,
                ControlFlow::ExitSub       => return ControlFlow::ExitSub,
                ControlFlow::ExitFunction  => return ControlFlow::ExitFunction,
                ControlFlow::ExitProperty  => return ControlFlow::ExitProperty,
                ControlFlow::ResumeCurrent => return ControlFlow::ResumeCurrent,
            }
            
            // Check condition at end
            match should_continue(ctx) {
                Ok(true) => { /* continue looping */ }
                Ok(false) => {
                    eprintln!("🔁 DO LOOP END: post-condition false");
                    break;
                }
                Err(flow) => return flow,
            }
        }
    }
    
    ControlFlow::Continue
}

fn eval_opt(expr: &Expression, ctx: &mut Context) -> Option<Value> {
    crate::interpreter::evaluate_expression(expr, ctx).ok()
}

fn is_truthy(v: &Value) -> bool {
    match v {
        Value::Boolean(b) => *b,
        Value::Integer(i) => *i != 0,
        Value::Long(i)        => *i != 0,
        Value::LongLong(i)    => *i != 0,
        Value::Object(None) => false,                 // Nothing => false
        Value::Object(Some(inner)) => is_truthy(inner), // delegate
        Value::Byte(b)    => *b != 0,
        Value::Currency(c) => *c != 0.0,
        Value::Date(_)    => true,
        Value::Double(f)  => *f != 0.0,
        Value::Decimal(f) => *f != 0.0,
        Value::Single(f) => *f != 0.0,              
        Value::String(s)  => !s.is_empty(),
        Value::Array(arr) => !arr.is_empty(),
        Value::UserType { .. } => true,
    }
}

fn to_string(v: &Value) -> String {
    match v {
        Value::Single(f) => f.to_string(),         
        Value::String(s)  => s.clone(),
        Value::Integer(i) => i.to_string(),
        Value::Long(i)      => i.to_string(),
        Value::LongLong(i)  => i.to_string(),
        Value::Object(None) => "Nothing".into(),    
        Value::Object(Some(inner)) => to_string(inner),
        Value::Byte(b)    => b.to_string(),
        Value::Currency(c) => format!("{:.4}", c),
        Value::Date(d) => d.format("%m/%d/%Y").to_string(),
        Value::Double(f)  => f.to_string(),
        Value::Decimal(f) => f.to_string(),
        Value::Boolean(b) => if *b { "True".into() } else { "False".into() },
        Value::Array(arr) => format!("Array({})", arr.len()),  
        Value::UserType { type_name, .. } => format!("<{} instance>", type_name),
    }
}

fn value_to_integer(value: &Value) -> Result<i64, String> {
    match value {
        Value::Integer(i) => Ok(*i),
        &Value::Long(l)      => Ok(l as i64),
        &Value::LongLong(ll)  => Ok(ll),
        Value::Object(Some(inner)) => value_to_integer(inner),
        Value::Object(None) => Err("Cannot convert Nothing to integer".to_string()),
        Value::Byte(b)    => Ok(*b as i64),
        Value::Currency(c) => Ok(*c as i64),
        Value::Date(d) => {
            let base = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
                .ok_or("Invalid base date".to_string())?;
            Ok((*d - base).num_days())
        },
        Value::Double(f)  => Ok(*f as i64),
        Value::Decimal(f) => Ok(*f as i64),
        Value::Single(f) => Ok(*f as i64),
        Value::String(s)  => s.parse::<i64>().map_err(|_| format!("Cannot convert '{}' to integer", s)),
        Value::Boolean(b) => Ok(if *b { -1 } else { 0 }),
        Value::Array(_) => Err("Cannot convert Array to integer".to_string()),
        Value::UserType { type_name, .. } => { 
            Err(format!("Cannot convert {} to integer", type_name))
        }
    }
}

// Error raising that arms Resume and uses PC
fn raise_runtime_error(ctx: &mut Context, number: i32, description: &str, current_pc: usize) -> ControlFlow {
    ctx.err = Some(ErrObject {
        number,
        description: description.into(),
        source: "YourInterp".into(),
    });

    match ctx.on_error_mode {
        OnErrorMode::ResumeNext => ControlFlow::ResumeNext,
        OnErrorMode::GoTo => {
            ctx.resume_valid = true;
            ctx.resume_pc = Some(current_pc);
            if let Some(lbl) = ctx.on_error_label.clone() {
                ControlFlow::GoToLabel(lbl)
            } else {
                // Defensive: no label set, behave like unhandled
                ControlFlow::ExitSub
            }
        }
        OnErrorMode::None => ControlFlow::ExitSub,
    }
}

fn execute_enum_statement(
    visibility: Option<&str>,
    name: &str,
    members: &[EnumMember],
    ctx: &mut Context
) -> ControlFlow {
    let mut enum_members = HashMap::new();
    let mut next_value: i64 = 0;

    for member in members {
        let value = if let Some(ref expr) = member.value {
            match evaluate_expression(expr, ctx) {
                Ok(Value::Integer(v)) => { next_value = v; v }
                Ok(Value::String(s)) => match s.parse::<i64>() {
                    Ok(v) => { next_value = v; v }
                    Err(_) => { ctx.log(&format!("Error: Enum member '{}' has invalid value '{}'", member.name, s)); next_value }
                },
                Ok(other) => {
                    ctx.log(&format!("Error: Enum member '{}' value not integer: {:?}", member.name, other));
                    next_value
                }
                Err(e) => {
                    ctx.log(&format!("Error: Could not evaluate enum member '{}' value: {}", member.name, e));
                    next_value
                }
            }
        } else {
            next_value
        };

        enum_members.insert(member.name.clone(), value);
        next_value = value + 1;
    }

    ctx.define_enum(name.to_string(), enum_members);
    let _visibility_str = visibility.unwrap_or("Private");
    ControlFlow::Continue

}
fn execute_type_statement(
    visibility: Option<&str>,
    name: &str,
    fields: &[TypeField],
    ctx: &mut Context
) -> ControlFlow {
    let mut type_fields = HashMap::new();
    
    for field in fields {
        let field_def = FieldDefinition {
            name: field.name.clone(),
            field_type: field.field_type.clone(),
            string_length: field.string_length,
            is_array: field.dimensions.is_some(),
        };
        
        type_fields.insert(field.name.clone(), field_def);
    }
    
    ctx.define_type(name.to_string(), type_fields.clone());
    
    let visibility_str = visibility.unwrap_or("Public");
    ctx.log(&format!(
        "Defined {} Type {} with {} fields",
        visibility_str, name, fields.len()
    ));
    
    ControlFlow::Continue
}
fn maybe_handle_error(ctx: &mut Context, pc: usize) -> Option<ControlFlow> {
    //println!("      🔍 maybe_handle_error called at pc={}", pc);
    //println!("      🔍 ctx.err = {:?}", ctx.err);
    //println!("      🔍 on_error_mode = {:?}", ctx.on_error_mode);
    if ctx.err.is_none() {
        //println!("      🔍 No error set, returning None");
        return None;
    }

    match ctx.on_error_mode {
        OnErrorMode::ResumeNext => {
            //println!("      🔍 ResumeNext mode: arming resume at pc={}", pc);
            
            // Arm a resume point at the faulting statement
            ctx.resume_valid = true;
            ctx.resume_pc = Some(pc);
            // DO NOT clear Err here; user code may read Err.Number/Description
            Some(ControlFlow::ResumeNext) // skip the faulting statement
        }

        OnErrorMode::GoTo => {
            //println!("      🔍 GoTo mode: arming resume at pc={}", pc);
            
            // Arm resume so handler can later `Resume` / `Resume Next`
            ctx.resume_valid = true;
            ctx.resume_pc = Some(pc);

            if let Some(ref dest) = ctx.on_error_label {
                //println!("      🔍 Jumping to label: {}", dest);
                
                // DO NOT clear Err here; handler often reads Err.*
                Some(ControlFlow::GoToLabel(dest.clone()))
            } else {
                // Defensive: “On Error GoTo” without a label → end proc
                //println!("      🔍 No label set, exiting sub");
                
                Some(ControlFlow::ExitSub)
            }
        }

        OnErrorMode::None => {
            //println!("      🔍 No error handling, exiting sub");
            
            // Unhandled runtime error → terminate this procedure (bubble if you prefer)
            Some(ControlFlow::ExitSub)
        }
    }
}
