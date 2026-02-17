use crate::ast::{Statement, ForStatement, DoWhileStatement, Expression, OnErrorKind, ResumeKind, EnumMember,TypeField, DoWhileConditionType};
use crate::interpreter::evaluate_expression;
use crate::context::{Context, Value, ScopeKind, FieldDefinition, ErrObject, OnErrorMode};
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
    ErrorGoToLabel(String),
    ResumeNext,      // On Error Resume Next, or Resume Next
    ResumeCurrent,
    FramePushed,   // Indicates a new frame was pushed, don't advance
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
        

        Statement::Comment(_text) => {
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

        // Record functions for later calls
        Statement::Function { name, params, return_type, body } => {
            ctx.define_function(name.clone(), params.clone(), body.clone(), return_type.clone());
            ctx.log(&format!("Defined function {}", name));
            ControlFlow::Continue
        }

        // Record Property Get for later calls
        Statement::PropertyGet { name, params, return_type, body } => {
            ctx.register_property("Get", name, params, body);
            // Store return type if needed
            if let Some(ref rt) = return_type {
                ctx.function_return_types.insert(format!("Get_{}", name), Some(rt.clone()));
            }
            ctx.log(&format!("Defined Property Get {}", name));
            ControlFlow::Continue
        }

        // Record Property Let for later calls
        Statement::PropertyLet { name, params, body } => {
            ctx.register_property("Let", name, params, body);
            ctx.log(&format!("Defined Property Let {}", name));
            ControlFlow::Continue
        }

        // Record Property Set for later calls
        Statement::PropertySet { name, params, body } => {
            ctx.register_property("Set", name, params, body);
            ctx.log(&format!("Defined Property Set {}", name));
            ControlFlow::Continue
        }

        // ReDim statement - resize arrays
        Statement::ReDim { preserve, variables } => {
            // TODO: Full array support not yet implemented
            // For now, just log the ReDim operation
            for var in variables {
                let bounds_str: Vec<String> = var.bounds.iter().map(|b| {
                    let lower_str = match &b.lower {
                        Some(_) => "N".to_string(),
                        None => "0".to_string(),
                    };
                    format!("{} To N", lower_str)
                }).collect();
                ctx.log(&format!("ReDim {}{} ({})", 
                    if *preserve { "Preserve " } else { "" },
                    var.name, 
                    bounds_str.join(", ")));
            }
            ControlFlow::Continue
        }

        // Dim declarations with declared-type defaults
        // Statement::Dim { names } => {
        //     for (v, maybe_type) in names {
        //         let initial_value = if let Some(type_name) = maybe_type {
        //             // First check if it's a user-defined type
        //             if ctx.is_type_defined(type_name) {
        //                 match ctx.create_type_instance(type_name) {
        //                     Some(instance) => {
        //                         // Set the type in context
        //                         ctx.set_var_type(v.clone(), crate::context::DeclaredType::Variant);
        //                         instance
        //                     }
        //                     None => {
        //                         // fallback to empty string
        //                         ctx.set_var_type(v.clone(), crate::context::DeclaredType::Variant);
        //                         Value::String(String::new())
        //                     }
        //                 }
        //             } else {
        //                 let ty = crate::context::DeclaredType::from_opt_str(Some(type_name));
        //                 ctx.set_var_type(v.clone(), ty);
        //                 match ty {
        //                     crate::context::DeclaredType::Byte     => Value::Byte(0),
        //                     crate::context::DeclaredType::Integer  => Value::Integer(0),
        //                     crate::context::DeclaredType::Long     => Value::Long(0),
        //                     crate::context::DeclaredType::LongLong => Value::LongLong(0),
        //                     crate::context::DeclaredType::Object   => Value::Object(None), 
        //                     crate::context::DeclaredType::Currency => Value::Currency(0.0),
        //                     crate::context::DeclaredType::Date     => chrono::NaiveDate::from_ymd_opt(1899,12,30).map(Value::Date).unwrap_or(Value::Date(chrono::NaiveDate::from_ymd_opt(1899,12,30).unwrap())),
        //                     crate::context::DeclaredType::Double   => Value::Double(0.0),
        //                     crate::context::DeclaredType::Decimal  => Value::Decimal(0.0),
        //                     crate::context::DeclaredType::Single   => Value::Single(0.0),
        //                     crate::context::DeclaredType::String   => Value::String(String::new()),
        //                     crate::context::DeclaredType::Boolean  => Value::Boolean(false),
        //                     crate::context::DeclaredType::Variant  => Value::String(String::new()),
        //                 }
        //             }
        //         } else {
        //             // No type specified - default to Variant
        //             ctx.set_var_type(v.clone(), crate::context::DeclaredType::Variant);
        //             Value::String(String::new())
        //         };
        //         ctx.declare_local(v.clone(), initial_value);
        //     }
        //     ControlFlow::Continue
        // }

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
                            crate::context::DeclaredType::Variant  => Value::Empty,  // Uninitialized Variant is Empty
                        }
                    }
                } else {
                    // No type specified - default to Variant (Empty)
                    ctx.set_var_type(v.clone(), crate::context::DeclaredType::Variant);
                    Value::Empty
                };
                ctx.declare_local(v.clone(), initial_value);
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
        // Statement::Assignment { lvalue, rvalue } => {
        //     // 1) Evaluate the RHS expression safely, catching interpreter errors
        //     let rhs_val_res = crate::interpreter::evaluate_expression(rvalue, ctx);
        
        //     if let Err(e) = rhs_val_res.as_ref() {
        //         // Capture the runtime error into the VBA Err object
        //         ctx.err = Some(ErrObject {
        //             number: 13, // VBA Type mismatch or general eval error
        //             description: e.to_string(),
        //             source: "Interpreter".into(),
        //         });
        //     }
        
        //     // 2) If any error is active, handle it according to On Error settings
        //     if ctx.err.is_some() {
        //         if let Some(flow) = maybe_handle_error(ctx, pc) {
        //             // Skip the assignment entirely when an error occurs
        //             return flow;
        //         }
        //         // In "On Error Resume Next" mode, skip assignment and continue
        //         return ControlFlow::Continue;
        //     }
        
        //     // 3) Safe unwrap – expression evaluated successfully
        //     let rhs_val = match rhs_val_res {
        //         Ok(v) => v,
        //         Err(_) => return ControlFlow::Continue, // safety fallback
        //     };
        
        //     // 4) Now perform the actual assignment
        //     match lvalue {
        //         // ————————————————————————————————————————————————————————————
        //         // PROPERTY ACCESS: e.g., Emp1.FirstName = "John"
        //         // ————————————————————————————————————————————————————————————
        //         crate::ast::AssignmentTarget::PropertyAccess { object, property } => {
        //             if let Some(mut obj_val) = ctx.get_var(object) {
        //                 // Coerce value if object has typed fields (optional)
        //                 match obj_val.set_field(property, rhs_val.clone()) {
        //                     Ok(()) => {
        //                         // Update the object in context
        //                         ctx.set_var(object.to_string(), obj_val);
        //                         ctx.log(&format!("Set {}.{} = {}", object, property, rhs_val.as_string()));
        //                     }
        //                     Err(e) => {
        //                         ctx.log(&format!("Error setting field: {}", e));
        //                         ctx.err = Some(ErrObject {
        //                             number: 13,
        //                             description: format!("Error setting field: {}", e),
        //                             source: "Interpreter".into(),
        //                         });
        //                         if let Some(flow) = maybe_handle_error(ctx, pc) {
        //                             return flow;
        //                         }
        //                         return ControlFlow::Continue;
        //                     }
        //                 }
        //             } else {
        //                 ctx.log(&format!("Error: Variable '{}' not found", object));
        //                 ctx.err = Some(ErrObject {
        //                     number: 91, // "Object variable or With block variable not set"
        //                     description: format!("Variable '{}' not found", object),
        //                     source: "Interpreter".into(),
        //                 });
        //                 if let Some(flow) = maybe_handle_error(ctx, pc) {
        //                     return flow;
        //                 }
        //                 return ControlFlow::Continue;
        //             }
        //         }
        
        //         // ————————————————————————————————————————————————————————————
        //         // IDENTIFIER ASSIGNMENT: e.g., x = 42 or Name = "Alice"
        //         // ————————————————————————————————————————————————————————————
        //         crate::ast::AssignmentTarget::Identifier(var_name) => {
        //             if let Some(ty) = ctx.get_var_type(var_name) {
        //                 // Coerce to declared type
        //                 match crate::interpreter::coerce::coerce_to_declared(rhs_val, ty) {
        //                     Ok(v) => {
        //                         ctx.set_var(var_name.clone(), v);
        //                     }
        //                     Err(e) => {
        //                         ctx.log(&format!("Type mismatch assigning to {}: {}", var_name, e));
        //                         ctx.err = Some(ErrObject {
        //                             number: 13,
        //                             description: format!("Type mismatch assigning to {}: {}", var_name, e),
        //                             source: "Interpreter".into(),
        //                         });
        //                         if let Some(flow) = maybe_handle_error(ctx, pc) {
        //                             return flow;
        //                         }
        //                         return ControlFlow::Continue;
        //                     }
        //                 }
        //             } else {
        //                 // No declared type => Variant semantics
        //                 ctx.set_var(var_name.clone(), rhs_val);
        //             }
        //         }
        //     }
        
        //     ControlFlow::Continue
        // }
        Statement::Assignment { lvalue, rvalue } => {
            let had_previous_error = ctx.err.is_some();
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
            // Only trigger error handling if this is a NEW error
            if ctx.err.is_some() && !had_previous_error {
                if let Some(flow) = maybe_handle_error(ctx, pc) {
                    return flow;
                }
            }

            // ✅ ONLY handle errors in GoTo mode if we just set resume_valid
            // In ResumeNextAuto mode, errors are already handled in evaluate_expression
            if ctx.err.is_some() && ctx.on_error_mode == OnErrorMode::GoTo {
                // Check if this is a FRESH error (resume_valid just became true)
                if ctx.resume_valid {
                    if let Some(flow) = maybe_handle_error(ctx, pc) {
                        return flow;
                    }
                }
            }
            
            // In Resume Next mode, just continue
            if ctx.err.is_some() && ctx.on_error_mode == OnErrorMode::ResumeNextAuto {
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
                    // Evaluate the object expression (supports Range("B" & i), Worksheets(...).Range(...), etc.)
                    // The object is now an Expression, so we can evaluate it properly
                    
                    // Handle WithMethodCall objects (e.g., .Range("A1").Value = xxx inside With block)
                    if let crate::ast::Expression::WithMethodCall { method, args } = object.as_ref() {
                        if method.eq_ignore_ascii_case("Range") {
                            // Get the With object (should be a Worksheet)
                            if let Some(_with_obj) = ctx.with_stack.last().cloned() {
                                // Evaluate the Range argument
                                if let Some(addr_expr) = args.first() {
                                    match crate::interpreter::evaluate_expression(addr_expr, ctx) {
                                        Ok(val) => {
                                            let address = match val {
                                                Value::String(s) => s,
                                                other => other.as_string(),
                                            };
                                            // Set the Range property
                                            match crate::host::excel::properties::set_property("range", &address, property, rhs_val.clone(), ctx) {
                                                Ok(_) => {
                                                    ctx.log(&format!("Set .Range(\"{}\").{} = {}", address, property, rhs_val.as_string()));
                                                    return ControlFlow::Continue;
                                                }
                                                Err(e) => {
                                                    ctx.err = Some(ErrObject {
                                                        number: 13,
                                                        description: format!("Error setting Range property: {}", e),
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
                                            ctx.err = Some(ErrObject {
                                                number: 11,
                                                description: e.to_string(),
                                                source: "Interpreter".into(),
                                            });
                                            if let Some(flow) = maybe_handle_error(ctx, pc) {
                                                return flow;
                                            }
                                            return ControlFlow::Continue;
                                        }
                                    }
                                }
                            } else {
                                ctx.err = Some(ErrObject {
                                    number: 91,
                                    description: "'.Range()' used outside of With block".to_string(),
                                    source: "Interpreter".into(),
                                });
                                if let Some(flow) = maybe_handle_error(ctx, pc) {
                                    return flow;
                                }
                                return ControlFlow::Continue;
                            }
                        }
                    }
                    
                    // Try to handle FunctionCall objects (e.g., Range(...).something)
                    if let crate::ast::Expression::FunctionCall { function, args } = object.as_ref() {
                        if let crate::ast::Expression::Identifier(fn_name) = function.as_ref() {
                            if fn_name.eq_ignore_ascii_case("Range") {
                                // Case: Range(...).Value = xxx
                                if let Some(arg) = args.first() {
                                    // Evaluate the argument (supports "B1", "B" & i, etc.)
                                    match crate::interpreter::evaluate_expression(arg, ctx) {
                                        Ok(val) => {
                                            let address = match val {
                                                Value::String(s) => s,
                                                other => other.as_string(),
                                            };
                                            match crate::host::excel::properties::set_property("range", &address, property, rhs_val.clone(), ctx) {
                                                Ok(_) => return ControlFlow::Continue,
                                                Err(e) => {
                                                    ctx.err = Some(ErrObject {
                                                        number: 13,
                                                        description: format!("Error setting Range property: {}", e),
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
                                            ctx.err = Some(ErrObject {
                                                number: 11,
                                                description: e.to_string(),
                                                source: "Interpreter".into(),
                                            });
                                            if let Some(flow) = maybe_handle_error(ctx, pc) {
                                                return flow;
                                            }
                                            return ControlFlow::Continue;
                                        }
                                    }
                                }
                                return ControlFlow::Continue;
                            }
                        }
                        // Handle PropertyAccess.PropertyAccess with FunctionCall (e.g., Worksheets(...).Range(...).Value)
                        else if let crate::ast::Expression::PropertyAccess { obj: inner_obj, property: inner_prop } = function.as_ref() {
                            if inner_prop.eq_ignore_ascii_case("Range") {
                                // We have Worksheets(...).Range(...).Value
                                // The Range(...) is the function call's function part, so we need to get the first arg of our function call
                                if let Some(range_arg) = args.first() {
                                    match crate::interpreter::evaluate_expression(range_arg, ctx) {
                                        Ok(val) => {
                                            let address = match val {
                                                Value::String(s) => s,
                                                other => other.as_string(),
                                            };
                                            match crate::host::excel::properties::set_property("range", &address, property, rhs_val.clone(), ctx) {
                                                Ok(_) => return ControlFlow::Continue,
                                                Err(e) => {
                                                    ctx.err = Some(ErrObject {
                                                        number: 13,
                                                        description: format!("Error setting Range property: {}", e),
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
                                            ctx.err = Some(ErrObject {
                                                number: 11,
                                                description: e.to_string(),
                                                source: "Interpreter".into(),
                                            });
                                            if let Some(flow) = maybe_handle_error(ctx, pc) {
                                                return flow;
                                            }
                                            return ControlFlow::Continue;
                                        }
                                    }
                                }
                                return ControlFlow::Continue;
                            }
                        }
                    }
                    
                    // Fallback: treat object as identifier
                    if let crate::ast::Expression::Identifier(obj_name) = object.as_ref() {
                        // Check if object variable is declared (Option Explicit)
                        if let Err(e) = ctx.validate_variable_usage(obj_name) {
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
                        
                        // ✨ NEW: Check for COM object property set
                        if ctx.com_registry.get_global(obj_name).is_some() {
                            match crate::host::dispatch_com_call(
                                obj_name,
                                property,
                                Some(&[rhs_val.clone()]),
                                true,  // Is a set
                                ctx,
                            ) {
                                Ok(_) => return ControlFlow::Continue,
                                Err(e) => {
                                    ctx.err = Some(ErrObject {
                                        number: 13, // Type mismatch, or more specific COM error
                                        description: format!("COM error: {}", e),
                                        source: "Interpreter".into(),
                                    });
                                    if let Some(flow) = maybe_handle_error(ctx, pc) {
                                        return flow;
                                    }
                                    return ControlFlow::Continue;
                                }
                            }
                        }
                        
                        if let Some(mut obj_val) = ctx.get_var(obj_name) {
                            match obj_val.set_field(property, rhs_val.clone()) {
                                Ok(()) => {
                                    ctx.set_var(obj_name.to_string(), obj_val);
                                    ctx.log(&format!("Set {}.{} = {}", obj_name, property, rhs_val.as_string()));
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
                            ctx.log(&format!("Error: Variable '{}' not found", obj_name));
                            ctx.err = Some(ErrObject {
                                number: 91,
                                description: format!("Variable '{}' not found", obj_name),
                                source: "Interpreter".into(),
                            });
                            if let Some(flow) = maybe_handle_error(ctx, pc) {
                                return flow;
                            }
                            return ControlFlow::Continue;
                        }
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

                crate::ast::AssignmentTarget::WithMemberAccess { property } => {
                    // Handle .Property = value inside a With block
                    if ctx.with_stack.is_empty() {
                        ctx.err = Some(ErrObject {
                            number: 91,
                            description: "Invalid use of '.' - no With object in scope".to_string(),
                            source: "Interpreter".into(),
                        });
                        if let Some(flow) = maybe_handle_error(ctx, pc) {
                            return flow;
                        }
                        return ControlFlow::Continue;
                    }
                    
                    // Get mutable reference to the last with object and set the field
                    let result = {
                        let with_obj = ctx.with_stack.last_mut().unwrap();
                        with_obj.set_field(property, rhs_val.clone())
                    };
                    
                    match result {
                        Ok(()) => {
                            ctx.log(&format!("Set With.{} = {}", property, rhs_val.as_string()));
                        }
                        Err(e) => {
                            let err_msg = format!("Error setting With field: {}", e);
                            ctx.log(&err_msg);
                            ctx.err = Some(ErrObject {
                                number: 13,
                                description: err_msg,
                                source: "Interpreter".into(),
                            });
                            if let Some(flow) = maybe_handle_error(ctx, pc) {
                                return flow;
                            }
                            return ControlFlow::Continue;
                        }
                    }
                }

                crate::ast::AssignmentTarget::WithMethodCall { method, args } => {
                    // Handle .Method(args).Property = value inside a With block (e.g., .Range("A1").Value = 5)
                    if ctx.with_stack.is_empty() {
                        ctx.err = Some(ErrObject {
                            number: 91,
                            description: "Invalid use of '.' - no With object in scope".to_string(),
                            source: "Interpreter".into(),
                        });
                        if let Some(flow) = maybe_handle_error(ctx, pc) {
                            return flow;
                        }
                        return ControlFlow::Continue;
                    }
                    
                    // Get the With object (should be a Worksheet)
                    let with_obj = ctx.with_stack.last().cloned();
                    
                    if let Some(Value::Object(Some(inner))) = with_obj {
                        if let Value::String(obj_str) = inner.as_ref() {
                            // Check if this is a Worksheet reference
                            if obj_str.to_lowercase().starts_with("worksheet:") {
                                let sheet_name = obj_str.strip_prefix("worksheet:").unwrap_or(obj_str);
                                
                                // If method is "Range", this is .Range("A1").Value = xxx
                                if method.eq_ignore_ascii_case("Range") {
                                    // Evaluate the Range argument
                                    if let Some(addr_expr) = args.first() {
                                        match crate::interpreter::evaluate_expression(addr_expr, ctx) {
                                            Ok(Value::String(addr)) => {
                                                // Set the Range property using the sheet context
                                                // For now, we'll use the sheet_name in the address
                                                match crate::host::excel::properties::set_property("range", &addr, "Value", rhs_val.clone(), ctx) {
                                                    Ok(_) => {
                                                        ctx.log(&format!("Set {}.Range(\"{}\").Value = {}", sheet_name, addr, rhs_val.as_string()));
                                                        return ControlFlow::Continue;
                                                    }
                                                    Err(e) => {
                                                        ctx.err = Some(ErrObject {
                                                            number: 13,
                                                            description: format!("Error setting Range property: {}", e),
                                                            source: "Interpreter".into(),
                                                        });
                                                        if let Some(flow) = maybe_handle_error(ctx, pc) {
                                                            return flow;
                                                        }
                                                        return ControlFlow::Continue;
                                                    }
                                                }
                                            }
                                            Ok(other) => {
                                                // Non-string argument - convert to string
                                                let addr = other.as_string();
                                                match crate::host::excel::properties::set_property("range", &addr, "Value", rhs_val.clone(), ctx) {
                                                    Ok(_) => {
                                                        ctx.log(&format!("Set {}.Range(\"{}\").Value = {}", sheet_name, addr, rhs_val.as_string()));
                                                        return ControlFlow::Continue;
                                                    }
                                                    Err(e) => {
                                                        ctx.err = Some(ErrObject {
                                                            number: 13,
                                                            description: format!("Error setting Range property: {}", e),
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
                                                ctx.err = Some(ErrObject {
                                                    number: 11,
                                                    description: e.to_string(),
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
                            }
                        }
                    }
                    
                    ctx.err = Some(ErrObject {
                        number: 438,
                        description: format!("Object doesn't support this property or method: .{}", method),
                        source: "Interpreter".into(),
                    });
                    if let Some(flow) = maybe_handle_error(ctx, pc) {
                        return flow;
                    }
                    return ControlFlow::Continue;
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

        Statement::With { object, body } => {
            // Evaluate the With object expression
            match crate::interpreter::evaluate_expression(object, ctx) {
                Ok(obj_value) => {
                    // Push the object onto the With stack
                    ctx.with_stack.push(obj_value);
                    
                    // Execute the body statements
                    let result = execute_statement_list(body, ctx);
                    
                    // Pop the object from the With stack
                    ctx.with_stack.pop();
                    
                    result
                }
                Err(e) => {
                    ctx.err = Some(ErrObject {
                        number: 91,
                        description: format!("With object evaluation failed: {}", e),
                        source: "Interpreter".into(),
                    });
                    if let Some(flow) = maybe_handle_error(ctx, pc) {
                        return flow;
                    }
                    ControlFlow::Continue
                }
            }
        }

        Statement::Exit(exit_type) => ControlFlow::from_exit_type(exit_type),

        Statement::Label(_) => ControlFlow::Continue,

        Statement::Expression(expr) => {
            let _ = eval_opt(expr, ctx);
            ControlFlow::Continue
        }

        // ——— Error handling directives
        Statement::OnError(kind) => {
            match kind {
                OnErrorKind::ResumeNext => { 
                    ctx.on_error_mode = OnErrorMode::ResumeNextAuto;  // ← CHANGED from ResumeNext
                    ctx.on_error_label = None;
                    ctx.resume_valid = false;
                    ctx.resume_location = None;
                }
                OnErrorKind::GoToLabel(lbl) => { 
                    ctx.on_error_mode = OnErrorMode::GoTo;
                    ctx.on_error_label = Some(lbl.clone());
                    ctx.resume_valid = false;
                    ctx.resume_location = None;
                }
                OnErrorKind::GoToZero => {
                    ctx.on_error_mode = OnErrorMode::None;
                    ctx.on_error_label = None;
                    ctx.resume_valid = false;
                    ctx.resume_location = None;
                }
            }
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
                ctx.declare_variable(&param.name);  // Use param.name for Parameter struct
                ctx.declare_local(param.name.clone(), val);
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
                | ControlFlow::ErrorGoToLabel(_)
                | ControlFlow::ResumeNext
                | ControlFlow::FramePushed
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
            ControlFlow::ErrorGoToLabel(lbl) => {
                // This is for the VM to handle, just bubble outward
                return ControlFlow::ErrorGoToLabel(lbl);
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
            | ControlFlow::FramePushed
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
            println!("🔁 FOR LOOP END: condition false (counter={})", counter);
            break;
        }
        println!("\n🔁 --- For iteration: {} = {} ---", for_stmt.counter, counter);
       
        match execute_statement_list(&for_stmt.body, ctx) {
            ControlFlow::Continue => { 
                //println!("🔁 Loop body returned Continue");
            /* keep looping */ }

            ControlFlow::ExitFor      => {
                println!("🔁 ExitFor encountered");
                return ControlFlow::Continue;
            }
            ControlFlow::ContinueFor  => {  /* step and re-check */ }

            ControlFlow::ResumeNext   => { /* already advanced by list */ }
            ControlFlow::GoToLabel(s) =>{
                println!("🔁 GoToLabel encountered: {}", s);
                 return ControlFlow::GoToLabel(s);}

            ControlFlow::ExitDo        => return ControlFlow::ExitDo,
            ControlFlow::ContinueDo    => return ControlFlow::ContinueDo,
            ControlFlow::ExitWhile     => return ControlFlow::ExitWhile,
            ControlFlow::ContinueWhile => return ControlFlow::ContinueWhile,
            ControlFlow::ExitSelect    => return ControlFlow::ExitSelect,

            ControlFlow::ExitSub       => return ControlFlow::ExitSub,
            ControlFlow::ExitFunction  => return ControlFlow::ExitFunction,
            ControlFlow::ErrorGoToLabel(lbl) => return ControlFlow::ErrorGoToLabel(lbl),
            ControlFlow::ExitProperty  => return ControlFlow::ExitProperty,

            ControlFlow::ResumeCurrent => return ControlFlow::ResumeCurrent,
            ControlFlow::FramePushed => return ControlFlow::FramePushed,
        }

        // Step
        counter += step_int;
        println!("🔁 Stepping: {} = {}", for_stmt.counter, counter);
        ctx.set_var(for_stmt.counter.clone(), Value::Integer(counter));
    }

    ControlFlow::Continue
}

pub fn execute_do_while_loop(do_stmt: &DoWhileStatement, ctx: &mut Context, pc: usize) -> ControlFlow {
    use DoWhileConditionType::*;
    
    // eprintln!("\n🔁 === DO WHILE LOOP START: type={:?}, test_at_end={} ===", 
            //   do_stmt.condition_type, do_stmt.test_at_end);
    
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
                    // eprintln!("🔁 DO LOOP END: pre-condition false");
                    break;
                }
                Err(flow) => return flow,
            }
            
            // eprintln!("🔁 --- Do loop iteration (pre-test) ---");
            
            match execute_statement_list(&do_stmt.body, ctx) {
                ControlFlow::Continue => { /* keep looping */ }
                
                ControlFlow::ExitDo => {
                    // eprintln!("🔁 ExitDo encountered");
                    return ControlFlow::Continue;
                }
                
                ControlFlow::ContinueDo => { /* continue to next iteration */ }
                
                ControlFlow::GoToLabel(s) => {
                    // eprintln!("🔁 GoToLabel encountered: {}", s);
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
                ControlFlow::ErrorGoToLabel(lbl) => return ControlFlow::ErrorGoToLabel(lbl),
                ControlFlow::ExitProperty  => return ControlFlow::ExitProperty,
                ControlFlow::ResumeCurrent => return ControlFlow::ResumeCurrent,
                ControlFlow::FramePushed => return ControlFlow::FramePushed,
            }
        }
    } 
    // Post-test loops: Do...Loop While/Until
    else {
        loop {
            // eprintln!("🔁 --- Do loop iteration (post-test) ---");
            
            match execute_statement_list(&do_stmt.body, ctx) {
                ControlFlow::Continue => { /* keep looping */ }
                
                ControlFlow::ExitDo => {
                    // eprintln!("🔁 ExitDo encountered");
                    return ControlFlow::Continue;
                }
                
                ControlFlow::ContinueDo => { /* check condition and continue */ }
                
                ControlFlow::GoToLabel(s) => {
                    // eprintln!("🔁 GoToLabel encountered: {}", s);
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
                ControlFlow::ErrorGoToLabel(lbl) => return ControlFlow::ErrorGoToLabel(lbl),
                ControlFlow::ExitProperty  => return ControlFlow::ExitProperty,
                ControlFlow::ResumeCurrent => return ControlFlow::ResumeCurrent,
                ControlFlow::FramePushed => return ControlFlow::FramePushed,
            }
            
            // Check condition at end
            match should_continue(ctx) {
                Ok(true) => { /* continue looping */ }
                Ok(false) => {
                    // eprintln!("🔁 DO LOOP END: post-condition false");
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
        Value::DateTime(_) => true,
        Value::Time(_) => true,
        Value::Double(f)  => *f != 0.0,
        Value::Decimal(f) => *f != 0.0,
        Value::Single(f) => *f != 0.0,              
        Value::String(s)  => !s.is_empty(),
        Value::UserType { .. } => true,
        Value::Empty => false,
        Value::Null => false,
        Value::Error(_) => false,  // Error values are falsy
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
        Value::DateTime(dt) => dt.format("%m/%d/%Y %H:%M:%S").to_string(),
        Value::Time(t) => t.format("%H:%M:%S").to_string(),
        Value::Double(f)  => f.to_string(),
        Value::Decimal(f) => f.to_string(),
        Value::Boolean(b) => if *b { "True".into() } else { "False".into() },
        Value::UserType { type_name, .. } => format!("<{} instance>", type_name),
        Value::Empty => String::new(),  
        Value::Null => "Null".into(),
        Value::Error(e) => format!("Error {}", e),
    }
}

pub fn value_to_integer(value: &Value) -> Result<i64, String> {
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
        Value::DateTime(dt) => {
            let base = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
                .ok_or("Invalid base date".to_string())?;
            Ok((dt.date() - base).num_days())
        },
        Value::Time(_) => Ok(0), // Time alone has no date component
        Value::Double(f)  => Ok(*f as i64),
        Value::Decimal(f) => Ok(*f as i64),
        Value::Single(f) => Ok(*f as i64),
        Value::String(s)  => s.parse::<i64>().map_err(|_| format!("Cannot convert '{}' to integer", s)),
        Value::Boolean(b) => Ok(if *b { -1 } else { 0 }),
        Value::UserType { type_name, .. } => { 
            Err(format!("Cannot convert {} to integer", type_name))
        }
        Value::Empty => Ok(0),
        Value::Null => Err("Cannot convert Null to integer".to_string()),
        Value::Error(e) => Ok(*e as i64),  // Error values convert to their error number
    }
}

// Error raising that arms Resume and uses PC
fn raise_runtime_error(
    ctx: &mut Context,
    number: i32,
    description: &str,
    current_pc: usize,
) -> ControlFlow {
    // eprintln!(
    //     "💥 raise_runtime_error: number={}, desc='{}', pc={}, mode={:?}",
    //     number, description, current_pc, ctx.on_error_mode
    // );

    ctx.err = Some(ErrObject {
        number,
        description: description.into(),
        source: "Interpreter".into(),
    });

    match ctx.on_error_mode {
        OnErrorMode::ResumeNextAuto => {
            // eprintln!("   → On Error Resume Next: skipping failing statement");
            ControlFlow::Continue  // Just continue to next statement
        }

        OnErrorMode::GoTo => {
            // eprintln!("   → On Error GoTo: arming resume state");
            ctx.resume_valid = true;
            ctx.resume_pc = Some(current_pc);
            
            if let Some(lbl) = ctx.on_error_label.clone() {
                // eprintln!("   → jumping to handler label '{}'", lbl);
                // eprintln!(
                //     "   [GoToLabel] err.is_some={} on_error_mode={:?} resume_valid={} resume_pc={:?}",
                //     ctx.err.is_some(),
                //     ctx.on_error_mode,
                //     ctx.resume_valid,
                //     ctx.resume_pc,
                // );
                ControlFlow::ErrorGoToLabel(lbl)
            } else {
                // eprintln!("   → No label set; exiting Sub");
                ControlFlow::ExitSub
            }
        }

        OnErrorMode::None => {
            // eprintln!("   → No error handler: exiting Sub");
            ControlFlow::ExitSub
        }
    }
}
fn maybe_handle_error(ctx: &mut Context, pc: usize) -> Option<ControlFlow> {
    if ctx.err.is_none() {
        return None;
    }

    match ctx.on_error_mode {
        OnErrorMode::ResumeNextAuto => {
            ctx.resume_valid = true;
            ctx.resume_pc = Some(pc);
            Some(ControlFlow::Continue)
        }

        OnErrorMode::GoTo => {
            ctx.resume_valid = true;
            ctx.resume_pc = Some(pc);
            // ✅ NEW: Store which frame the error occurred in
            // We'll use a new field in Context for this
            // ctx.error_frame_id = Some(current_frame_id);
            // But we don't have frame_id here... so we need another approach

            if let Some(ref dest) = ctx.on_error_label {
                Some(ControlFlow::ErrorGoToLabel(dest.clone()))
            } else {
                Some(ControlFlow::ExitSub)
            }
        }

        OnErrorMode::None => {
            Some(ControlFlow::ExitSub)
        }
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
