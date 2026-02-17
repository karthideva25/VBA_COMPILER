use anyhow::{anyhow, bail, Result};
use crate::ast::Expression;
use crate::context::{Context, Value};
use crate::interpreter::coerce;
use super::builtins::{resolve_builtin_identifier};

pub(crate) fn evaluate_expression(expr: &Expression, ctx: &mut Context) -> Result<Value> {
    use Expression::*;

    match expr {
        // ——— Literals
        Integer(n) => Ok(Value::Integer(*n)),
        Byte(b) => Ok(Value::Byte(*b)),
        String(s)  => Ok(Value::String(s.clone())),
        Boolean(b) => Ok(Value::Boolean(*b)),
        Double(f)  => Ok(Value::Double(*f)),
        Decimal(f) => Ok(Value::Decimal(*f)),
        Single(s) => Ok(Value::Single(*s)),
        Currency(c) => Ok(Value::Currency(*c)),
        Date(d)     => Ok(Value::Date(*d)),

        // ——— Identifiers: built-in constants first, then variables
        Identifier(name) => {
            // 1. Check built-in constants first (vbTrue, vbCrLf, etc.)
            if let Some(v) = resolve_builtin_identifier(name) {
                return Ok(v);
            }
            
            // 2. Check if it's a qualified enum reference (EnumName.MemberName)
            //    This handles cases where the parser stored it as a single identifier
            if let Some(value) = ctx.resolve_enum_member(name) {
                return Ok(value);
            }
            
            // 3. Check Option Explicit before reading variable
            if let Err(e) = ctx.validate_variable_usage(name) {
                return Err(anyhow::anyhow!("{}", e));
            }
            
            // 4. Regular variable lookup
            ctx.get_var(name)
                .ok_or_else(|| anyhow::anyhow!("Variable '{}' not found", name))
        }
        
        BuiltInConstant(name) => {
            resolve_builtin_identifier(name)
                .ok_or_else(|| anyhow::anyhow!("Unknown built-in constant: {}", name))
        }

        // ——— Unary: op is a String (e.g., "+", "-", "Not")
        UnaryOp { op, expr } => {
            let v = evaluate_expression(expr, ctx)?;
            crate::interpreter::operations::eval_unary(op.as_str(), v)
        }

        // ——— Binary: op is a String (e.g., "+", "*", "<>", etc.)
        BinaryOp { left: lhs, op, right: rhs } => {
            // Evaluate children first
            let l = evaluate_expression(lhs, ctx)?;
            let r = evaluate_expression(rhs, ctx)?;
        
            // Delegate to pure ops
            crate::interpreter::operations::eval_binary(ctx, op, l, r)
        }

        // ——— Function calls used as expressions
        // FunctionCall { function, args } => {
        //     // Only simple identifier calls supported for now
        //     let name = if let Expression::Identifier(n) = &**function {
        //         n
        //     } else {
        //         bail!("Only simple identifier calls supported for now")
        //     };
        //      // Try builtin functions first
        //     if let Ok(Some(val)) = crate::interpreter::builtins::functions::handle_builtin_call(name, args, ctx) {
        //         return Ok(val);
        //     }
        
        //     // Built-in: Format(value, pattern)
        //     if name.eq_ignore_ascii_case("Format") {
        //         if args.len() != 2 {
        //             bail!("Format expects 2 arguments");
        //         }
        //         let v = evaluate_expression(&args[0], ctx)?;
        //         let p = evaluate_expression(&args[1], ctx)?;
        //         let pat = match p {
        //             Value::String(s) => s,
        //             other => bail!("Format second argument must be a string, got {:?}", other),
        //         };
        
        //         // VBA→chrono pattern adapter
        //         let mut chrono_pat = pat.clone();
        //         chrono_pat = chrono_pat.replace("yyyy", "%Y");
        //         chrono_pat = chrono_pat.replace("mm", "%m");
        //         chrono_pat = chrono_pat.replace("dd", "%d");
        //         chrono_pat = chrono_pat.replace("hh", "%H");
        //         chrono_pat = chrono_pat.replace("ss", "%S");
        
        //         let s = match v {
        //             Value::Date(d) => {
        //                 let dt = d.and_hms_opt(0, 0, 0).expect("valid midnight");
        //                 dt.format(&chrono_pat).to_string()
        //             }
        //             _ => crate::interpreter::coerce::to_string(&v),
        //         };
        //         return Ok(Value::String(s));
        //     }
        
        //     // Unknown functions -> for now, just return 0
        //     let _evaluated: Vec<Value> =
        //         args.iter().map(|a| evaluate_expression(a, ctx)).collect::<Result<_>>()?;
        //     Ok(Value::Integer(0))
        // }        


        FunctionCall { function, args } => {
            let name = if let Expression::Identifier(n) = &**function {
                n
            } else {
                bail!("Only simple identifier calls supported for now")
            };
            
            // Check if this is array access (variable exists and is an array)
            if let Some(var_value) = ctx.get_var(name) {
                if let Value::Array(arr) = var_value {
                    // This is array element access
                    if args.len() == 1 {
                        let index_val = evaluate_expression(&args[0], ctx)?;
                        let index = coerce::to_i64(&index_val)? as usize;
                        
                        // VBA arrays are 1-based
                        let zero_based_index = if index > 0 { index - 1 } else { 0 };
                        
                        if zero_based_index < arr.len() {
                            return Ok(arr[zero_based_index].clone());
                        } else {
                            bail!("Array index {} out of bounds (length {})", index, arr.len());
                        }
                    } else {
                        bail!("Multi-dimensional array access not yet supported");
                    }
                }
            }
            
            // Built-in: Format(value, pattern)
            if name.eq_ignore_ascii_case("Format") {
                if args.len() != 2 {
                    bail!("Format expects 2 arguments");
                }
                let v = evaluate_expression(&args[0], ctx)?;
                let p = evaluate_expression(&args[1], ctx)?;
                let pat = match p {
                    Value::String(s) => s,
                    other => bail!("Format second argument must be a string, got {:?}", other),
                };

                let mut chrono_pat = pat.clone();
                chrono_pat = chrono_pat.replace("yyyy", "%Y");
                chrono_pat = chrono_pat.replace("mm", "%m");
                chrono_pat = chrono_pat.replace("dd", "%d");
                chrono_pat = chrono_pat.replace("hh", "%H");
                chrono_pat = chrono_pat.replace("ss", "%S");

                let s = match v {
                    Value::Date(d) => {
                        let dt = d.and_hms_opt(0, 0, 0).expect("valid midnight");
                        dt.format(&chrono_pat).to_string()
                    }
                    _ => crate::interpreter::coerce::to_string(&v),
                };
                return Ok(Value::String(s));
            }

            // Unknown functions -> return 0
            let _evaluated: Vec<Value> =
                args.iter().map(|a| evaluate_expression(a, ctx)).collect::<Result<_>>()?;
            Ok(Value::Integer(0))
        }
        // ——— Property Access: Handle enum member access and user types
        PropertyAccess { obj, property } => {
            // 1) Evaluate the object expression first
            let object_val = evaluate_expression(obj, ctx)?;
        
            // 2) Handle user-defined types (Type ... End Type)
            if let Value::UserType { fields, type_name } = object_val {
                if let Some(val) = fields.get(property) {
                    return Ok(val.clone());
                } else {
                    bail!("Field '{}' not found on type '{}'", property, type_name);
                }
            }
        
            // 3) Handle enum member access (EnumName.Member)
            if let Expression::Identifier(enum_name) = &**obj {
                if let Some(value) = ctx.get_enum_value(enum_name, property) {
                    return Ok(Value::Integer(value));
                }
            }
        
            // 4) Handle special-case VBA Err object properties
            if let Expression::Identifier(name) = &**obj {
                if name.eq_ignore_ascii_case("Err") {
                    match property.to_ascii_lowercase().as_str() {
                        "number" => {
                            let n = ctx.err.as_ref().map(|e| e.number).unwrap_or(0);
                            return Ok(Value::Integer(n.into()));
                        }
                        "description" => {
                            let d = ctx.err.as_ref()
                                .map(|e| e.description.clone())
                                .unwrap_or_default();
                            return Ok(Value::String(d));
                        }
                        "clear" => {
                            // VBA Err.Clear is a subroutine (no return)
                            ctx.err = None;
                            ctx.resume_valid = false;
                            return Ok(Value::Integer(0));
                        }
                        "source" => {
                            let s = ctx.err.as_ref()
                                .map(|e| e.source.clone())
                                .unwrap_or_default();
                            return Ok(Value::String(s));
                        }
                        _ => bail!("Unknown Err property: {}", property),
                    }
                }
            }
        
            // 5) Fallback: if we reach here, property access type was unsupported
            match object_val {
                Value::String(_) | Value::Integer(_) | Value::Boolean(_) => {
                    bail!(
                        "Property access '{:?}.{}' not supported for {:?}",
                        obj, property, object_val
                    )
                }
                other => bail!("Cannot access property '{}' on {:?}", property, other),
            }
        }
    }
}