use crate::ast::{Expression, Statement, ForStatement, ExitType};
use crate::context::{Context, Value, ScopeKind};

// Enhanced control flow enum
#[derive(Debug, Clone, PartialEq)]
enum ControlFlow {
    Continue,
    ExitFor,
    ExitDo,
    ExitSub,
    ExitFunction,
    ExitProperty,
}

pub fn execute_ast(node: &Statement, ctx: &mut Context) {
    let _ = execute_statement(node, ctx);
}

fn execute_statement(stmt: &Statement, ctx: &mut Context) -> ControlFlow {
    match stmt {
        Statement::BlankLine => { ControlFlow::Continue}

        Statement::Comment(text) => {
            println!("Comment: {}", text);
             ControlFlow::Continue
        }
        // Record subroutines for later calls (unchanged public behavior)
        Statement::Subroutine { name, params, body } => {
            ctx.define_sub(name.clone(), params.clone(), body.clone());
            ctx.log(&format!("Defined subroutine {}", name));
             ControlFlow::Continue
        }

        // DIM declares a variable in the *current* scope.
        // If we're at top level (no scope), it falls back to global (per Context::declare_local).
        Statement::Dim { names } => {
            for (v, _maybe_type) in names {
                ctx.declare_local(v.clone(), Value::String(String::new()));
            }
             ControlFlow::Continue
        }

        // SET/Assignment: scope-aware write via Context::set_var (updates nearest declared var, else global)
        Statement::Set { target, expr } => {
            if let Some(val) = evaluate_expression(expr, ctx) {
                ctx.set_var(target.clone(), val);
            }
             ControlFlow::Continue
        }

        Statement::Assignment { lvalue, rvalue } => {
            if let Some(val) = evaluate_expression(rvalue, ctx) {
                ctx.set_var(lvalue.clone(), val);
            }
             ControlFlow::Continue

        }

        Statement::MsgBox { expr } => {
            if let Some(val) = evaluate_expression(expr, ctx) {
                ctx.log(&val.as_string());
            }
             ControlFlow::Continue
        }

        Statement::GoTo { label } => {
            ctx.log(&format!("*** Goto `{}` not implemented", label));
            ControlFlow::Continue
        }

        Statement::If { condition, then_branch, else_if, else_branch } => {
            if let Some(cv) = evaluate_expression(condition, ctx) {
                if is_truthy(&cv) {
                    // THEN
                    println!("Executing Then branch");
                    for s in then_branch {
                        let _ = execute_statement(s, ctx);
                    }
                    return ControlFlow::Continue; // short-circuit: never consider ElseIf/Else
                }
                // ELSEIFs
            for (elseif_condition, elseif_statements) in else_if.iter() {
                if let Some(elseif_value) = evaluate_expression(elseif_condition, ctx) {
                    if is_truthy(&elseif_value) {
                        println!("Executing ElseIf branch");
                        for stmt in elseif_statements {
                            let _ = execute_statement(stmt, ctx);
                        }
                        return ControlFlow::Continue; // short-circuit on first true ElseIf
                    }
                }
            }

            // ELSE
                if !else_branch.is_empty() {
                    println!("Executing Else branch");
                    for s in else_branch {
                        execute_statement(s, ctx);
                    }
                }
            }
             ControlFlow::Continue
        }
        Statement::For(for_stmt) => {
            execute_for_loop(for_stmt, ctx)
        }

         Statement::Exit(exit_type) => {
            ctx.log(&format!("Exit {}", exit_type.as_str()));
            ControlFlow::from_exit_type(exit_type)
        }

        Statement::Label(_) => {ControlFlow::Continue}

        Statement::Expression(expr) => {
            let _ = evaluate_expression(expr, ctx);
            ControlFlow::Continue
        }

        // === Subroutine call with proper scoping ===
        Statement::Call { function, args } => {
             // 1) Handle built-ins first (MsgBox, etc.)
             if handle_builtin_call(function, args, ctx) {
                return ControlFlow::Continue;
            }

            // Clone the (params, body) out of `ctx.subs` to avoid borrow conflicts before we mutate `ctx`
            let (params, body) = match ctx.subs.get(function).cloned() {
                Some(pb) => pb,
                None => {
                    ctx.log(&format!("*** Call `{}` not implemented", function));
                    return ControlFlow::Continue;
                }
            };

            // Arity check
            if params.len() != args.len() {
                ctx.log(&format!(
                    "*** Error: Sub `{}` expects {} arguments, got {}",
                    function,
                    params.len(),
                    args.len()
                ));
                return ControlFlow::Continue;
            }

            // Evaluate arguments *in the caller context* before entering callee scope
            let mut arg_vals = Vec::with_capacity(args.len());
            for (idx, a) in args.iter().enumerate() {
                match evaluate_expression(a, ctx) {
                    Some(v) => arg_vals.push(v),
                    None => {
                        ctx.log(&format!(
                            "*** Error: could not evaluate argument {} for `{}`",
                            idx + 1,
                            function
                        ));
                        return ControlFlow::Continue;
                    }
                }
            }

            ctx.log(&format!("Entering Sub {}", function));

            // New procedure scope
            ctx.push_scope(function.clone(), ScopeKind::Subroutine);

            // Bind parameters as locals (so they shadow globals)
            for (param, val) in params.iter().zip(arg_vals.into_iter()) {
                ctx.declare_local(param.clone(), val);
            }

            // Execute the sub body inside the new scope
            // for s in &body {
            //     execute_statement(s, ctx);
            // }

            // // Tear down procedure scope
            // ctx.pop_scope();

            // ctx.log(&format!("Leaving Sub {}", function));

            // Execute sub body with control flow handling
            let flow = execute_statement_list(&body, ctx);
            
            ctx.pop_scope();
            ctx.log(&format!("Leaving Sub {}", function));

            // Exit Sub only affects the current sub, not the caller
            if flow == ControlFlow::ExitSub {
                ControlFlow::Continue
            } else {
                flow
            }

            
        }
    }
}

/// Handle intrinsic/built-in calls that may appear as `Statement::Call`.
/// Returns true if the call was handled.
fn handle_builtin_call(function: &str, args: &[Expression], ctx: &mut Context) -> bool {
    // MSGBOX — allow statement-style and call-style; ignore extra args for now
    if function.eq_ignore_ascii_case("MsgBox") || function.eq_ignore_ascii_case("Msgbox") {
        let text = args.get(0)
            .and_then(|e| evaluate_expression(e, ctx))
            .map(|v| v.as_string())
            .unwrap_or_default();
        ctx.log(&text);
        return true;
    }

    // You can extend this with more built-ins:
    // if function.eq_ignore_ascii_case("InputBox") { ... }
    // if function.eq_ignore_ascii_case("Print") || function.eq_ignore_ascii_case("Debug.Print") { ... }

    false
}
fn is_truthy(v: &Value) -> bool {
    match v {   
        Value::Integer(i) => *i != 0,
        Value::Byte(b)    => *b != 0, 
        Value::String(s) => {
            let low = s.to_lowercase();
            !low.is_empty() && low != "false"
        }
        Value::Boolean(b) => *b,
    }
    
}
fn evaluate_expression(expr: &Expression, ctx: &Context) -> Option<Value> {
    //println!(" Evaluating expression:"); // Debug print
    match expr {
        Expression::Integer(i) => Some(Value::Integer(*i)),
        Expression::Byte(b) => Some(Value::Byte(*b)),
        Expression::String(s) => Some(Value::String(s.clone())),
        Expression::Identifier(n) => ctx.get_var(n),
        Expression::Boolean(b) => Some(Value::Boolean(*b)),
        Expression::BinaryOp { left, op, right } => {
            let l = evaluate_expression(left, ctx)?;
            let r = evaluate_expression(right, ctx)?;
            
            match (l.clone(), r.clone()) {
                //boolean operation
                (Value::Boolean(b1), Value::Boolean(b2)) => match op.as_str() {
                    "And" => Some(Value::Boolean(b1 && b2)),
                    "Or"  => Some(Value::Boolean(b1 || b2)),
                    _ => None,
                },

                // Integer operations
                (Value::Integer(li), Value::Integer(ri)) => match op.as_str() {
                    // Arithmetic operators
                    "+" => Some(Value::Integer(li + ri)),
                    "-" => Some(Value::Integer(li - ri)),
                    "*" => Some(Value::Integer(li * ri)),
                    "/" => {
                        if ri == 0 {
                            println!("❌ Division by zero error");
                            None
                        } else {
                            Some(Value::Integer(li / ri))
                        }
                    }
                    // String concatenation with integers
                    "&" => Some(Value::String(li.to_string() + &ri.to_string())),
                    
                    // Comparison operators (return Integer: -1 for True, 0 for False in VBA)
                    "=" => Some(Value::Integer(if li == ri { -1 } else { 0 })),
                    "<>" => Some(Value::Integer(if li != ri { -1 } else { 0 })),
                    ">" => Some(Value::Integer(if li > ri { -1 } else { 0 })),
                    "<" => Some(Value::Integer(if li < ri { -1 } else { 0 })),
                    ">=" => Some(Value::Integer(if li >= ri { -1 } else { 0 })),
                    "<=" => Some(Value::Integer(if li <= ri { -1 } else { 0 })),
                    
                    _ => {
                        println!("⚠️ Unsupported integer operation: {}", op);
                        None
                    }
                },
                
                // Binary operations: Byte vs Byte (promote to Integer if necessary)
                (Value::Byte(lb), Value::Byte(rb)) => match op.as_str() {
                    "+" => Some(Value::Integer(lb as i64 + rb as i64)),
                    "-" => Some(Value::Integer(lb as i64 - rb as i64)),
                    "*" => Some(Value::Integer(lb as i64 * rb as i64)),
                    "/" => {
                        if rb == 0 {
                            println!("❌ Division by zero");
                            None
                        } else {
                            Some(Value::Integer(lb as i64 / rb as i64))
                        }
                    },
                    "&" => Some(Value::String(lb.to_string() + &rb.to_string())),
                    "=" => Some(Value::Integer(if lb == rb { -1 } else { 0 })),
                    "<>" => Some(Value::Integer(if lb != rb { -1 } else { 0 })),
                    ">" => Some(Value::Integer(if lb > rb { -1 } else { 0 })),
                    "<" => Some(Value::Integer(if lb < rb { -1 } else { 0 })),
                    ">=" => Some(Value::Integer(if lb >= rb { -1 } else { 0 })),
                    "<=" => Some(Value::Integer(if lb <= rb { -1 } else { 0 })),
                    _ => None,
                },

                // Mixed Byte and Integer
                (Value::Byte(lb), Value::Integer(ri)) => match op.as_str() {
                    "+" => Some(Value::Integer(lb as i64 + ri)),
                    "-" => Some(Value::Integer(lb as i64 - ri)),
                    "*" => Some(Value::Integer(lb as i64 * ri)),
                    "/" => Some(Value::Integer(lb as i64 / ri)),
                    "&" => Some(Value::String(lb.to_string() + &ri.to_string())),
                    _ => None,
                },
                (Value::Integer(li), Value::Byte(rb)) => match op.as_str() {
                    "+" => Some(Value::Integer(li + rb as i64)),
                    "-" => Some(Value::Integer(li - rb as i64)),
                    "*" => Some(Value::Integer(li * rb as i64)),
                    "/" => Some(Value::Integer(li / rb as i64)),
                    "&" => Some(Value::String(li.to_string() + &rb.to_string())),
                    _ => None,
                },

                // String operations
                (Value::String(ls), Value::String(rs)) => match op.as_str() {
                    // String concatenation
                    "&" => Some(Value::String(ls + &rs)),
                    
                    // String comparison
                    "=" => Some(Value::Integer(if ls == rs { -1 } else { 0 })),
                    "<>" => Some(Value::Integer(if ls != rs { -1 } else { 0 })),
                    ">" => Some(Value::Integer(if ls > rs { -1 } else { 0 })),
                    "<" => Some(Value::Integer(if ls < rs { -1 } else { 0 })),
                    ">=" => Some(Value::Integer(if ls >= rs { -1 } else { 0 })),
                    "<=" => Some(Value::Integer(if ls <= rs { -1 } else { 0 })),
                    
                    // Handle incorrect + for string concatenation (backward compatibility)
                    "+" => {
                        println!("⚠️ Warning: '+' used for string concatenation, should be '&'");
                        Some(Value::String(ls + &rs))
                    }
                    
                    _ => {
                        println!("⚠️ Unsupported string operation: {}", op);
                        None
                    }
                },
                
                // Mixed String & Integer operations
                (Value::String(ls), Value::Integer(ri)) => match op.as_str() {
                    "&" => Some(Value::String(ls + &ri.to_string())),
                    "+" => {
                        println!("⚠️ Warning: '+' used for string concatenation, should be '&'");
                        Some(Value::String(ls + &ri.to_string()))
                    }
                    // String vs Integer comparison (convert integer to string for comparison)
                    "=" => Some(Value::Integer(if ls == ri.to_string() { -1 } else { 0 })),
                    "<>" => Some(Value::Integer(if ls != ri.to_string() { -1 } else { 0 })),
                    _ => {
                        println!("⚠️ Unsupported string-integer operation: {}", op);
                        None
                    }
                },
                
                // Mixed Integer & String operations  
                (Value::Integer(li), Value::String(rs)) => match op.as_str() {
                    "&" => Some(Value::String(li.to_string() + &rs)),
                    "+" => {
                        println!("⚠️ Warning: '+' used for string concatenation, should be '&'");
                        Some(Value::String(li.to_string() + &rs))
                    }
                    // Integer vs String comparison (convert integer to string for comparison)
                    "=" => Some(Value::Integer(if li.to_string() == rs { -1 } else { 0 })),
                    "<>" => Some(Value::Integer(if li.to_string() != rs { -1 } else { 0 })),
                    _ => {
                        println!("⚠️ Unsupported integer-string operation: {}", op);
                        None
                    }
                },               
                 _ => todo!(),
                
            }
        }
        Expression::UnaryOp { op, expr } => {
            match op.as_str() {
                "-" => {
                    evaluate_expression(expr, ctx).and_then(|val| {
                        match val {
                            Value::Integer(i) => Some(Value::Integer(-i)),
                            Value::Byte(b)    => Some(Value::Integer(-(b as i64))),
                            Value::String(s) => {
                                // Try to convert string to integer for negation
                                s.parse::<i64>()
                                    .map(|i| Value::Integer(-i))
                                    .ok()
                            },
                             Value::Boolean(b) => {
                                // In VBA, True = -1, False = 0
                                // Negating a boolean flips its numeric equivalent
                                let v = if b { 1 } else { 0 };
                                Some(Value::Integer(-v))
                            }

                        }
                    })
                }
                "+" => {
                    // Unary plus - just return the value unchanged
                    evaluate_expression(expr, ctx)
                }
                "Not" | "not" => {
                    // Logical NOT - return -1 for True, 0 for False in VBA
                    evaluate_expression(expr, ctx).map(|val| {
                        let is_truthy = match val {
                            Value::Integer(i) => i != 0,
                            Value::Byte(b)    => b != 0,
                            Value::String(s) => !s.is_empty() && s.to_lowercase() != "false",
                            Value::Boolean(b) => b,
                        };
                        Value::Integer(if is_truthy { 0 } else { -1 })
                    })
                }
                _ => {
                    println!("⚠️ Unsupported unary operation: {}", op);
                    None
                }
            }
        }
        Expression::BuiltInConstant(name) => {
            let val = match name.as_str() {
                "vbCalGreg" => Value::Integer( 0),
                "vbCalHijri" => Value::Integer( 1),

                // CallType constants
                "vbMethod" => Value::Integer( -1),
                "vbGet" => Value::Integer( -2),
                "vbLet" => Value::Integer( -4),
                "vbSet" => Value::Integer( -8),

                 // Color constants
                "vbBlack" => Value::Integer( 0),
                "vbRed" => Value::Integer( 255),
                "vbGreen" => Value::Integer( 65280),
                "vbYellow" => Value::Integer( 65535),
                "vbBlue" => Value::Integer( 16711680),
                "vbMagenta" => Value::Integer( 16711935),
                "vbCyan" => Value::Integer( 16776960),
                "vbWhite" => Value::Integer( 16777215),

                 // Comparison constants
                "vbUseCompareOption"=> Value::Integer(-1),
                "vbBinaryCompare" => Value::Integer( 0),
                "vbTextCompare" => Value::Integer( 1),
                "vbDatabaseCompare" => Value::Integer( 2),

                // Day of Week constants
                "vbSunday" => Value::Integer( 1),
                "vbMonday" => Value::Integer( 2),
                "vbTuesday" => Value::Integer( 3),
                "vbWednesday" => Value::Integer( 4),
                "vbThursday" => Value::Integer( 5),
                "vbFriday" => Value::Integer( 6),
                "vbSaturday" => Value::Integer( 7),
                "vbUseSystemDayOfWeek" => Value::Integer( 0),

                // First Week of Year constants
                "vbUseSystem" => Value::Integer( 0),
                "vbFirstJan1" => Value::Integer( 1),
                "vbFirstFourDays" => Value::Integer( 2),
                "vbFirstFullWeek" => Value::Integer( 3),

                // Date/Time format constants
                "vbGeneralDate" => Value::Integer( 0),
                "vbLongDate" => Value::Integer( 1),
                "vbShortDate" => Value::Integer( 2),
                "vbLongTime" => Value::Integer( 3),
                "vbShortTime" => Value::Integer( 4),

                // Key Code Constants - Mouse Buttons
                "vbKeyLButton" => Value::Integer( 1),        // 0x1 - Left mouse button
                "vbKeyRButton" => Value::Integer( 2),        // 0x2 - Right mouse button
                "vbKeyCancel" => Value::Integer( 3),         // 0x3 - CANCEL key
                "vbKeyMButton" => Value::Integer( 4),        // 0x4 - Middle mouse button
                
                // Key Code Constants - Special Keys
                "vbKeyBack" => Value::Integer( 8),           // 0x8 - BACKSPACE key
                "vbKeyTab" => Value::Integer( 9),            // 0x9 - TAB key
                "vbKeyClear" => Value::Integer( 12),         // 0xC - CLEAR key
                "vbKeyReturn" => Value::Integer( 13),        // 0xD - ENTER key
                "vbKeyShift" => Value::Integer( 16),         // 0x10 - SHIFT key
                "vbKeyControl" => Value::Integer( 17),       // 0x11 - CTRL key
                "vbKeyMenu" => Value::Integer( 18),          // 0x12 - MENU key
                "vbKeyPause" => Value::Integer( 19),         // 0x13 - PAUSE key
                "vbKeyCapital" => Value::Integer( 20),       // 0x14 - CAPS LOCK key
                "vbKeyEscape" => Value::Integer( 27),        // 0x1B - ESC key
                "vbKeySpace" => Value::Integer( 32),         // 0x20 - SPACEBAR key
                
                // Key Code Constants - Navigation Keys
                "vbKeyPageUp" => Value::Integer( 33),        // 0x21 - PAGE UP key
                "vbKeyPageDown" => Value::Integer( 34),      // 0x22 - PAGE DOWN key
                "vbKeyEnd" => Value::Integer( 35),           // 0x23 - END key
                "vbKeyHome" => Value::Integer( 36),          // 0x24 - HOME key
                "vbKeyLeft" => Value::Integer( 37),          // 0x25 - LEFT ARROW key
                "vbKeyUp" => Value::Integer( 38),            // 0x26 - UP ARROW key
                "vbKeyRight" => Value::Integer( 39),         // 0x27 - RIGHT ARROW key
                "vbKeyDown" => Value::Integer( 40),          // 0x28 - DOWN ARROW key
                "vbKeySelect" => Value::Integer( 41),        // 0x29 - SELECT key
                "vbKeyPrint" => Value::Integer( 42),         // 0x2A - PRINT SCREEN key
                "vbKeyExecute" => Value::Integer( 43),       // 0x2B - EXECUTE key
                "vbKeySnapshot" => Value::Integer( 44),      // 0x2C - SNAPSHOT key
                "vbKeyInsert" => Value::Integer( 45),        // 0x2D - INSERT key
                "vbKeyDelete" => Value::Integer( 46),        // 0x2E - DELETE key
                "vbKeyHelp" => Value::Integer( 47),          // 0x2F - HELP key
                "vbKeyNumlock" => Value::Integer( 144),      // 0x90 - NUM LOCK key
                
                // Key Code Constants - Alphabetic Keys (A-Z)
                "vbKeyA" => Value::Integer( 65),             // ASCII 'A'
                "vbKeyB" => Value::Integer( 66),             // ASCII 'B'
                "vbKeyC" => Value::Integer( 67),             // ASCII 'C'
                "vbKeyD" => Value::Integer( 68),             // ASCII 'D'
                "vbKeyE" => Value::Integer( 69),             // ASCII 'E'
                "vbKeyF" => Value::Integer( 70),             // ASCII 'F'
                "vbKeyG" => Value::Integer( 71),             // ASCII 'G'
                "vbKeyH" => Value::Integer( 72),             // ASCII 'H'
                "vbKeyI" => Value::Integer( 73),             // ASCII 'I'
                "vbKeyJ" => Value::Integer( 74),             // ASCII 'J'
                "vbKeyK" => Value::Integer( 75),             // ASCII 'K'
                "vbKeyL" => Value::Integer( 76),             // ASCII 'L'
                "vbKeyM" => Value::Integer( 77),             // ASCII 'M'
                "vbKeyN" => Value::Integer( 78),             // ASCII 'N'
                "vbKeyO" => Value::Integer( 79),             // ASCII 'O'
                "vbKeyP" => Value::Integer( 80),             // ASCII 'P'
                "vbKeyQ" => Value::Integer( 81),             // ASCII 'Q'
                "vbKeyR" => Value::Integer( 82),             // ASCII 'R'
                "vbKeyS" => Value::Integer( 83),             // ASCII 'S'
                "vbKeyT" => Value::Integer( 84),             // ASCII 'T'
                "vbKeyU" => Value::Integer( 85),             // ASCII 'U'
                "vbKeyV" => Value::Integer( 86),             // ASCII 'V'
                "vbKeyW" => Value::Integer( 87),             // ASCII 'W'
                "vbKeyX" => Value::Integer( 88),             // ASCII 'X'
                "vbKeyY" => Value::Integer( 89),             // ASCII 'Y'
                "vbKeyZ" => Value::Integer( 90),             // ASCII 'Z'
                
                // Key Code Constants - Numeric Keys (0-9)
                "vbKey0" => Value::Integer( 48),             // ASCII '0'
                "vbKey1" => Value::Integer( 49),             // ASCII '1'
                "vbKey2" => Value::Integer( 50),             // ASCII '2'
                "vbKey3" => Value::Integer( 51),             // ASCII '3'
                "vbKey4" => Value::Integer( 52),             // ASCII '4'
                "vbKey5" => Value::Integer( 53),             // ASCII '5'
                "vbKey6" => Value::Integer( 54),             // ASCII '6'
                "vbKey7" => Value::Integer( 55),             // ASCII '7'
                "vbKey8" => Value::Integer( 56),             // ASCII '8'
                "vbKey9" => Value::Integer( 57),             // ASCII '9'
                
                // Key Code Constants - Numpad Keys
                "vbKeyNumpad0" => Value::Integer( 96),       // Numpad 0
                "vbKeyNumpad1" => Value::Integer( 97),       // Numpad 1
                "vbKeyNumpad2" => Value::Integer( 98),       // Numpad 2
                "vbKeyNumpad3" => Value::Integer( 99),       // Numpad 3
                "vbKeyNumpad4" => Value::Integer( 100),      // Numpad 4
                "vbKeyNumpad5" => Value::Integer( 101),      // Numpad 5
                "vbKeyNumpad6" => Value::Integer( 102),      // Numpad 6
                "vbKeyNumpad7" => Value::Integer( 103),      // Numpad 7
                "vbKeyNumpad8" => Value::Integer( 104),      // Numpad 8
                "vbKeyNumpad9" => Value::Integer( 105),      // Numpad 9
                "vbKeyMultiply" => Value::Integer( 106),     // Numpad * (multiply)
                "vbKeyAdd" => Value::Integer( 107),          // Numpad + (add)
                "vbKeySeparator" => Value::Integer( 108),    // Numpad separator
                "vbKeySubtract" => Value::Integer( 109),     // Numpad - (subtract)
                "vbKeyDecimal" => Value::Integer( 110),      // Numpad . (decimal)
                "vbKeyDivide" => Value::Integer( 111),       // Numpad / (divide)
                
                // Key Code Constants - Function Keys (F1-F16)
                "vbKeyF1" => Value::Integer( 112),           // F1 key
                "vbKeyF2" => Value::Integer( 113),           // F2 key
                "vbKeyF3" => Value::Integer( 114),           // F3 key
                "vbKeyF4" => Value::Integer( 115),           // F4 key
                "vbKeyF5" => Value::Integer( 116),           // F5 key
                "vbKeyF6" => Value::Integer( 117),           // F6 key
                "vbKeyF7" => Value::Integer( 118),           // F7 key
                "vbKeyF8" => Value::Integer( 119),           // F8 key
                "vbKeyF9" => Value::Integer( 120),           // F9 key
                "vbKeyF10" => Value::Integer( 121),          // F10 key
                "vbKeyF11" => Value::Integer( 122),          // F11 key
                "vbKeyF12" => Value::Integer( 123),          // F12 key
                "vbKeyF13" => Value::Integer( 124),          // F13 key
                "vbKeyF14" => Value::Integer( 125),          // F14 key
                "vbKeyF15" => Value::Integer( 126),          // F15 key
                "vbKeyF16" => Value::Integer( 127),          // F16 key

                
                // MsgBox constants
                "vbOKOnly" => Value::Integer( 0),
                "vbOKCancel" => Value::Integer( 1),
                "vbAbortRetryIgnore" => Value::Integer( 2),
                "vbYesNoCancel" => Value::Integer( 3),
                "vbYesNo" => Value::Integer( 4),
                "vbRetryCancel" => Value::Integer( 5),

                // MsgBox icon constants
                "vbCritical" => Value::Integer( 16),
                "vbQuestion" => Value::Integer( 32),
                "vbExclamation" => Value::Integer( 48),
                "vbInformation" => Value::Integer( 64),

                // MsgBox return values
                "vbOK" => Value::Integer( 1),
                "vbCancel" => Value::Integer( 2),
                "vbAbort" => Value::Integer( 3),
                "vbRetry" => Value::Integer( 4),
                "vbIgnore" => Value::Integer( 5),
                "vbYes" => Value::Integer( 6),
                "vbNo" => Value::Integer( 7),

                 // Text case constants
                "vbUpperCase"   => Value::Integer( 1),
                "vbLowerCase"   => Value::Integer( 2),
                "vbProperCase"  => Value::Integer( 3),

                // String width and script constants
                "vbWide"        => Value::Integer( 4),
                "vbNarrow"      => Value::Integer( 8),
                "vbKatakana"    => Value::Integer( 16),
                "vbHiragana"    => Value::Integer( 32),

                // Unicode constants
                "vbUnicode"     => Value::Integer( -1),
                "vbFromUnicode" => Value::Integer( -2),

                "vbTrue" => Value::Integer( -1),
                "vbFalse" => Value::Integer( 0),
                "vbUseDefault" => Value::Integer( 2),

               

                "vbEmpty" => Value::Integer( 0),
                "vbNull" => Value::Integer( 1),
                "vbInteger" => Value::Integer( 2),
                "vbLong" => Value::Integer( 3),
                "vbSingle" => Value::Integer( 4),
                "vbDouble" => Value::Integer( 5),
                "vbCurrency" => Value::Integer( 6),
                "vbDate" => Value::Integer( 7),
                "vbString" => Value::Integer( 8),
                "vbObject" => Value::Integer( 9),
                "vbError" => Value::Integer( 10),
                "vbBoolean" => Value::Integer( 11),
                "vbVariant" => Value::Integer( 12),
                "vbDataObject" => Value::Integer( 13),
                "vbDecimal" => Value::Integer( 14),
                "vbByte" => Value::Integer( 17),
                "vbUserDefinedType" => Value::Integer( 36),
                "vbArray" => Value::Integer( 8192),



                "vbCrLf"       => Value::String( "\r\n".to_string()),
                "vbCr"         => Value::String( "\r".to_string()),
                "vbLf"         => Value::String( "\n".to_string()),
                "vbNewLine"    => Value::String( "\n".to_string()),       // same as vbLf in many contexts
                "vbNullChar"   => Value::String( '\0'.to_string()),       // null character
                "vbNullString" => Value::String( "".to_string()),         // empty string
                "vbObjectError"=> Value::Integer( -2147221504), // typical COM error base (example value)
                "vbTab"        => Value::String( "\t".to_string()),
                "vbBack"       => Value::String( '\x08'.to_string()),     // backspace character
                "vbFormFeed"   => Value::String( '\x0C'.to_string()),     // form feed character
                "vbVerticalTab"=> Value::String( '\x0B'.to_string()),     // vertical tab character


                _ => {
                    println!("⚠️ Unknown builtin constant: {}", name);
                    return None;
                }
            };
            Some(val)
        },
        
        Expression::FunctionCall { function, args } => {
            if let Expression::Identifier(ref f) = **function {
                if f.eq_ignore_ascii_case("cstr") && args.len() == 1 {
                    return evaluate_expression(&args[0], ctx)
                        .map(|v| Value::String(v.as_string()));
                }
                // Optional: allow MsgBox in an *expression* context (returns vbOK = 1).
                // We can't log here because `evaluate_expression` only has `&Context`.
                if f.eq_ignore_ascii_case("msgbox") {
                    if let Some(_prompt) = args.get(0).and_then(|e| evaluate_expression(e, ctx)) {
                        return Some(Value::Integer(1)); // vbOK
                    }
                }
            }
            None
        }
        
        Expression::PropertyAccess { obj, .. } => {
            let _ = evaluate_expression(obj, ctx)?;
            None
        }
    }
}

fn execute_statement_list(statements: &[Statement], ctx: &mut Context) -> ControlFlow {
    for stmt in statements {
        let flow = execute_statement(stmt, ctx);
        if flow != ControlFlow::Continue {
            return flow;
        }
    }
    ControlFlow::Continue
}
fn execute_for_loop(for_stmt: &ForStatement, ctx: &mut Context) -> ControlFlow {
    // ... your existing setup code for start_int, end_int, step_int ...
    
    let start_val = match evaluate_expression(&for_stmt.start, ctx) {
        Some(val) => val,
        None => {
            ctx.log("Error: Could not evaluate start expression in for loop");
            return ControlFlow::Continue;
        }
    };
    
    let end_val = match evaluate_expression(&for_stmt.end, ctx) {
        Some(val) => val,
        None => {
            ctx.log("Error: Could not evaluate end expression in for loop");
            return ControlFlow::Continue;
        }
    };
    
    let step_val = if let Some(step_expr) = &for_stmt.step {
        match evaluate_expression(step_expr, ctx) {
            Some(val) => val,
            None => {
                ctx.log("Error: Could not evaluate step expression in for loop");
                return ControlFlow::Continue;
            }
        }
    } else {
        Value::Integer(1)
    };
    
    let start_int = match value_to_integer(&start_val) {
        Ok(val) => val,
        Err(msg) => {
            ctx.log(&format!("Error in for loop start value: {}", msg));
            return ControlFlow::Continue;
        }
    };
    
    let end_int = match value_to_integer(&end_val) {
        Ok(val) => val,
        Err(msg) => {
            ctx.log(&format!("Error in for loop end value: {}", msg));
            return ControlFlow::Continue;
        }
    };
    
    let step_int = match value_to_integer(&step_val) {
        Ok(val) => val,
        Err(msg) => {
            ctx.log(&format!("Error in for loop step value: {}", msg));
            return ControlFlow::Continue;
        }
    };
    
    if step_int == 0 {
        ctx.log("Error: For loop step cannot be zero");
        return ControlFlow::Continue;
    }
    
    ctx.log(&format!("Starting For loop: {} = {} To {} Step {}", 
             for_stmt.counter, start_int, end_int, step_int));
    
    let mut counter = start_int;
    ctx.set_var(for_stmt.counter.clone(), Value::Integer(counter));
    
    // Execute the loop with comprehensive Exit support
    loop {
        let should_continue = if step_int > 0 {
            counter <= end_int
        } else {
            counter >= end_int
        };
        
        if !should_continue {
            break;
        }
        
        // Execute loop body with control flow handling
        let flow = execute_statement_list(&for_stmt.body, ctx);
        
        match flow {
            ControlFlow::Continue => {
                // Continue with loop
            }
            ControlFlow::ExitFor => {
                ctx.log(&format!("For loop exited early at {} = {}", for_stmt.counter, counter));
                return ControlFlow::Continue;
            }
            ControlFlow::ExitDo => {
                // Exit Do doesn't affect For loops - continue
                ctx.log("Warning: Exit Do inside For loop has no effect");
            }
            flow if flow.should_exit_procedure() => {
                // Exit Sub/Function/Property should propagate out of the loop
                return flow;
            }
            _ => {
                // Other control flows continue the loop
            }
        }
        
        counter += step_int;
        ctx.set_var(for_stmt.counter.clone(), Value::Integer(counter));
    }
    
    ctx.log(&format!("For loop completed. Final {} = {}", for_stmt.counter, counter));
    ControlFlow::Continue
}


fn value_to_integer(value: &Value) -> Result<i64, String> {
    match value {
        Value::Integer(i) => Ok(*i),
        Value::Byte(b)    => Ok(*b as i64),
        Value::String(s) => {
            s.parse::<i64>()
                .map_err(|_| format!("Cannot convert '{}' to integer", s))
        },
        Value::Boolean(b) => {
            // VBA semantics: True = -1, False = 0
            Ok(if *b { -1 } else { 0 })
        }
    }
}
impl ControlFlow {
    fn from_exit_type(exit_type: &ExitType) -> Self {
        match exit_type {
            ExitType::For => ControlFlow::ExitFor,
            ExitType::Do => ControlFlow::ExitDo,
            ExitType::Sub => ControlFlow::ExitSub,
            ExitType::Function => ControlFlow::ExitFunction,
            ExitType::Property => ControlFlow::ExitProperty,
        }
    }
    
    fn should_exit_loop(&self) -> bool {
        matches!(self, ControlFlow::ExitFor | ControlFlow::ExitDo)
    }
    
    fn should_exit_procedure(&self) -> bool {
        matches!(self, ControlFlow::ExitSub | ControlFlow::ExitFunction | ControlFlow::ExitProperty)
    }
}