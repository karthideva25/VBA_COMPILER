use anyhow::Result;
use crate::ast::Expression;
use crate::context::{Context, Value};
use crate::interpreter::evaluate_expression;

/// Return Ok(Some(Value)) if handled; Ok(None) to let caller try user-defined subs/funcs.
pub(crate) fn handle_builtin_call(function: &str, args: &[Expression], ctx: &mut Context)
    -> Result<Option<Value>>
{
    match function.to_ascii_lowercase().as_str() {
        // MSGBOX — allow statement-style and call-style
        "msgbox" => {
            let text = if let Some(e0) = args.get(0) {
                let v = evaluate_expression(e0, ctx)?;
                value_to_string(&v)
            } else {
                String::new()
            };
            ctx.log(&text);
            Ok(Some(Value::Integer(1))) // VBA MsgBox returns button code
        }

        // LEN — returns length of string
        "len" => {
            if args.len() != 1 {
                ctx.log("*** Error: Len() expects 1 argument");
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => Ok(Some(Value::Integer(s.len() as i64))),
                _ => {
                    ctx.log("*** Error: Len() expects string argument");
                    Ok(Some(Value::Integer(0)))
                }
            }
        }

        // MID — Mid(string, start, [length])
        // VBA uses 1-based indexing
        "mid" => {
            if args.len() < 2 || args.len() > 3 {
                ctx.log("*** Error: Mid() expects 2 or 3 arguments");
                return Ok(Some(Value::String(String::new())));
            }

            let string_val = evaluate_expression(&args[0], ctx)?;
            let start_val = evaluate_expression(&args[1], ctx)?;
            
            let string = match string_val {
                Value::String(s) => s,
                _ => {
                    ctx.log("*** Error: Mid() first argument must be string");
                    return Ok(Some(Value::String(String::new())));
                }
            };

            let start = match start_val {
                Value::Integer(i) => i,
                _ => {
                    ctx.log("*** Error: Mid() second argument must be integer");
                    return Ok(Some(Value::String(String::new())));
                }
            };

            // VBA uses 1-based indexing
            let start_idx = ((start - 1).max(0) as usize).min(string.len());

            if args.len() == 3 {
                // Mid with length
                let length_val = evaluate_expression(&args[2], ctx)?;
                let length = match length_val {
                    Value::Integer(i) => i.max(0) as usize,
                    _ => {
                        ctx.log("*** Error: Mid() third argument must be integer");
                        return Ok(Some(Value::String(String::new())));
                    }
                };
                
                let end_idx = (start_idx + length).min(string.len());
                Ok(Some(Value::String(string[start_idx..end_idx].to_string())))
            } else {
                // Mid without length - return rest of string
                Ok(Some(Value::String(string[start_idx..].to_string())))
            }
        }

        // LEFT — Left(string, length)
        "left" => {
            if args.len() != 2 {
                ctx.log("*** Error: Left() expects 2 arguments");
                return Ok(Some(Value::String(String::new())));
            }

            let string_val = evaluate_expression(&args[0], ctx)?;
            let length_val = evaluate_expression(&args[1], ctx)?;

            match (string_val, length_val) {
                (Value::String(s), Value::Integer(len)) => {
                    let len = len.max(0) as usize;
                    let end = len.min(s.len());
                    Ok(Some(Value::String(s[..end].to_string())))
                }
                _ => {
                    ctx.log("*** Error: Left() type mismatch");
                    Ok(Some(Value::String(String::new())))
                }
            }
        }

        // RIGHT — Right(string, length)
        "right" => {
            if args.len() != 2 {
                ctx.log("*** Error: Right() expects 2 arguments");
                return Ok(Some(Value::String(String::new())));
            }

            let string_val = evaluate_expression(&args[0], ctx)?;
            let length_val = evaluate_expression(&args[1], ctx)?;

            match (string_val, length_val) {
                (Value::String(s), Value::Integer(len)) => {
                    let len = len.max(0) as usize;
                    let start = s.len().saturating_sub(len);
                    Ok(Some(Value::String(s[start..].to_string())))
                }
                _ => {
                    ctx.log("*** Error: Right() type mismatch");
                    Ok(Some(Value::String(String::new())))
                }
            }
        }

        // UCASE — UCase(string)
        "ucase" => {
            if args.len() != 1 {
                ctx.log("*** Error: UCase() expects 1 argument");
                return Ok(Some(Value::String(String::new())));
            }

            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => Ok(Some(Value::String(s.to_uppercase()))),
                _ => {
                    ctx.log("*** Error: UCase() expects string argument");
                    Ok(Some(Value::String(String::new())))
                }
            }
        }

        // LCASE — LCase(string)
        "lcase" => {
            if args.len() != 1 {
                ctx.log("*** Error: LCase() expects 1 argument");
                return Ok(Some(Value::String(String::new())));
            }

            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => Ok(Some(Value::String(s.to_lowercase()))),
                _ => {
                    ctx.log("*** Error: LCase() expects string argument");
                    Ok(Some(Value::String(String::new())))
                }
            }
        }

        // TRIM — Trim(string)
        "trim" => {
            if args.len() != 1 {
                ctx.log("*** Error: Trim() expects 1 argument");
                return Ok(Some(Value::String(String::new())));
            }

            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => Ok(Some(Value::String(s.trim().to_string()))),
                _ => {
                    ctx.log("*** Error: Trim() expects string argument");
                    Ok(Some(Value::String(String::new())))
                }
            }
        }

        // INSTR — InStr([start,] string1, string2)
        // Returns position of first occurrence (1-based), or 0 if not found
        "instr" => {
            if args.len() < 2 || args.len() > 3 {
                ctx.log("*** Error: InStr() expects 2 or 3 arguments");
                return Ok(Some(Value::Integer(0)));
            }

            let (start, str1, str2) = if args.len() == 3 {
                let start_val = evaluate_expression(&args[0], ctx)?;
                let str1_val = evaluate_expression(&args[1], ctx)?;
                let str2_val = evaluate_expression(&args[2], ctx)?;
                
                match (start_val, str1_val, str2_val) {
                    (Value::Integer(s), Value::String(s1), Value::String(s2)) => (s, s1, s2),
                    _ => {
                        ctx.log("*** Error: InStr() type mismatch");
                        return Ok(Some(Value::Integer(0)));
                    }
                }
            } else {
                let str1_val = evaluate_expression(&args[0], ctx)?;
                let str2_val = evaluate_expression(&args[1], ctx)?;
                
                match (str1_val, str2_val) {
                    (Value::String(s1), Value::String(s2)) => (1, s1, s2),
                    _ => {
                        ctx.log("*** Error: InStr() type mismatch");
                        return Ok(Some(Value::Integer(0)));
                    }
                }
            };

            // VBA uses 1-based indexing
            let start_idx = ((start - 1).max(0) as usize).min(str1.len());
            
            match str1[start_idx..].find(&str2) {
                Some(pos) => Ok(Some(Value::Integer((start_idx + pos + 1) as i64))),
                None => Ok(Some(Value::Integer(0)))
            }
        }

        _ => Ok(None), // Not a builtin, let caller handle it
    }
}

/// Convenience wrapper to match your original `-> bool` usage
pub(crate) fn handle_builtin_call_bool(function: &str, args: &[Expression], ctx: &mut Context) -> bool {
    matches!(handle_builtin_call(function, args, ctx), Ok(Some(_)))
}

fn value_to_string(v: &Value) -> String {
    match v {
        Value::Single(f) => f.to_string(), 
        Value::String(s)  => s.clone(),
        Value::Integer(i) => i.to_string(),
        Value::Long(l)      => l.to_string(),
        Value::LongLong(ll)  => ll.to_string(),
        Value::Object(None) => "Nothing".into(),
        Value::Object(Some(inner)) => value_to_string(inner),
        Value::Byte(b)    => b.to_string(),
        Value::Currency(c) => format!("{:.4}", c),
        Value::Date(d) => d.format("%Y-%m-%d").to_string(),
        Value::Double(f)  => f.to_string(),
        Value::Decimal(f) => f.to_string(),
        Value::Boolean(b) => if *b { "True".into() } else { "False".into() },
        Value::Array(arr) => {
             format!("Array({})", arr.len())  // CORRECT - returns String
        }
        Value::UserType { type_name, .. } => { 
            format!("<{} instance>", type_name)
        }
    }
}