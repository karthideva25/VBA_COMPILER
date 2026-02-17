//! VBA Math Functions
//! 
//! This module contains all VBA mathematical functions including:
//! - Abs, Sgn, Int, Fix
//! - Sqr, Exp, Log
//! - Sin, Cos, Tan, Atn
//! - Round, Rnd
//! - Hex, Oct

use anyhow::Result;
use crate::ast::Expression;
use crate::context::{Context, Value};
use crate::interpreter::evaluate_expression;

/// Handle math-related builtin function calls
pub(crate) fn handle_math_function(function: &str, args: &[Expression], ctx: &mut Context) -> Result<Option<Value>> {
    match function {
        // ============================================================
        // BASIC MATH
        // ============================================================

        // ABS — Returns absolute value
        "abs" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::Integer(i) => Ok(Some(Value::Integer(i.abs()))),
                Value::Long(l) => Ok(Some(Value::Long(l.abs()))),
                Value::Double(d) => Ok(Some(Value::Double(d.abs()))),
                Value::Single(s) => Ok(Some(Value::Single(s.abs()))),
                Value::Currency(c) => Ok(Some(Value::Currency(c.abs()))),
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // SGN — Returns sign of a number (-1, 0, or 1)
        "sgn" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = match &val {
                Value::Integer(i) => *i as f64,
                Value::Long(l) => *l as f64,
                Value::LongLong(ll) => *ll as f64,
                Value::Double(d) => *d,
                Value::Single(s) => *s as f64,
                _ => 0.0,
            };
            let sign = if f > 0.0 { 1 } else if f < 0.0 { -1 } else { 0 };
            Ok(Some(Value::Integer(sign)))
        }

        // INT — Returns integer portion (rounds toward negative infinity)
        "int" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = match &val {
                Value::Integer(i) => *i as f64,
                Value::Long(l) => *l as f64,
                Value::Double(d) => *d,
                Value::Single(s) => *s as f64,
                Value::Currency(c) => *c,
                _ => 0.0,
            };
            Ok(Some(Value::Integer(f.floor() as i64)))
        }

        // FIX — Returns integer portion (truncates toward zero)
        "fix" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = match &val {
                Value::Integer(i) => *i as f64,
                Value::Long(l) => *l as f64,
                Value::Double(d) => *d,
                Value::Single(s) => *s as f64,
                Value::Currency(c) => *c,
                _ => 0.0,
            };
            Ok(Some(Value::Integer(f.trunc() as i64)))
        }

        // ============================================================
        // ADVANCED MATH
        // ============================================================

        // SQR — Returns square root
        "sqr" => {
            if args.is_empty() {
                return Ok(Some(Value::Double(0.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = match &val {
                Value::Integer(i) => *i as f64,
                Value::Long(l) => *l as f64,
                Value::Double(d) => *d,
                Value::Single(s) => *s as f64,
                Value::Currency(c) => *c,
                _ => 0.0,
            };
            if f >= 0.0 {
                Ok(Some(Value::Double(f.sqrt())))
            } else {
                ctx.log("*** Error: Sqr() cannot take square root of negative number");
                Ok(Some(Value::Double(f64::NAN)))
            }
        }

        // EXP — Returns e raised to a power
        "exp" => {
            if args.is_empty() {
                return Ok(Some(Value::Double(1.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = match &val {
                Value::Integer(i) => *i as f64,
                Value::Long(l) => *l as f64,
                Value::Double(d) => *d,
                Value::Single(s) => *s as f64,
                _ => 0.0,
            };
            Ok(Some(Value::Double(f.exp())))
        }

        // LOG — Returns natural logarithm
        "log" => {
            if args.is_empty() {
                return Ok(Some(Value::Double(0.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = match &val {
                Value::Integer(i) => *i as f64,
                Value::Long(l) => *l as f64,
                Value::Double(d) => *d,
                Value::Single(s) => *s as f64,
                _ => 0.0,
            };
            if f > 0.0 {
                Ok(Some(Value::Double(f.ln())))
            } else {
                ctx.log("*** Error: Log() requires positive argument");
                Ok(Some(Value::Double(f64::NAN)))
            }
        }

        // ============================================================
        // TRIGONOMETRIC FUNCTIONS
        // ============================================================

        // SIN — Returns sine
        "sin" => {
            if args.is_empty() {
                return Ok(Some(Value::Double(0.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = match &val {
                Value::Integer(i) => *i as f64,
                Value::Long(l) => *l as f64,
                Value::Double(d) => *d,
                Value::Single(s) => *s as f64,
                _ => 0.0,
            };
            Ok(Some(Value::Double(f.sin())))
        }

        // COS — Returns cosine
        "cos" => {
            if args.is_empty() {
                return Ok(Some(Value::Double(1.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = match &val {
                Value::Integer(i) => *i as f64,
                Value::Long(l) => *l as f64,
                Value::Double(d) => *d,
                Value::Single(s) => *s as f64,
                _ => 0.0,
            };
            Ok(Some(Value::Double(f.cos())))
        }

        // TAN — Returns tangent
        "tan" => {
            if args.is_empty() {
                return Ok(Some(Value::Double(0.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = match &val {
                Value::Integer(i) => *i as f64,
                Value::Long(l) => *l as f64,
                Value::Double(d) => *d,
                Value::Single(s) => *s as f64,
                _ => 0.0,
            };
            Ok(Some(Value::Double(f.tan())))
        }

        // ATN — Returns arctangent
        "atn" => {
            if args.is_empty() {
                return Ok(Some(Value::Double(0.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = match &val {
                Value::Integer(i) => *i as f64,
                Value::Long(l) => *l as f64,
                Value::Double(d) => *d,
                Value::Single(s) => *s as f64,
                _ => 0.0,
            };
            Ok(Some(Value::Double(f.atan())))
        }

        // ============================================================
        // ROUNDING
        // ============================================================

        // ROUND — Rounds to specified decimal places (banker's rounding)
        "round" => {
            if args.is_empty() {
                return Ok(Some(Value::Double(0.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let decimals = if args.len() > 1 {
                match evaluate_expression(&args[1], ctx)? {
                    Value::Integer(i) => i,
                    _ => 0
                }
            } else { 0 };
            
            let f = match &val {
                Value::Integer(i) => *i as f64,
                Value::Long(l) => *l as f64,
                Value::Double(d) => *d,
                Value::Single(s) => *s as f64,
                Value::Currency(c) => *c,
                _ => 0.0,
            };
            
            let factor = 10_f64.powi(decimals as i32);
            let rounded = (f * factor).round() / factor;
            Ok(Some(Value::Double(rounded)))
        }

        // RND — Returns random number between 0 and 1
        "rnd" => {
            // Simple pseudo-random - in real VBA this uses a seed
            use std::time::{SystemTime, UNIX_EPOCH};
            let seed = SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap()
                .subsec_nanos();
            let random = (seed as f64 / u32::MAX as f64).fract();
            Ok(Some(Value::Single(random as f32)))
        }

        // ============================================================
        // NUMBER CONVERSION
        // ============================================================

        // HEX — Returns hexadecimal string representation
        "hex" | "hex$" => {
            if args.is_empty() {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let n = match val {
                Value::Integer(i) => i,
                Value::Long(l) => l as i64,
                Value::Byte(b) => b as i64,
                _ => 0
            };
            Ok(Some(Value::String(format!("{:X}", n))))
        }

        // OCT — Returns octal string representation
        "oct" | "oct$" => {
            if args.is_empty() {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let n = match val {
                Value::Integer(i) => i,
                Value::Long(l) => l as i64,
                Value::Byte(b) => b as i64,
                _ => 0
            };
            Ok(Some(Value::String(format!("{:o}", n))))
        }

        // VAL — Returns numeric value of a string
        "val" => {
            if args.is_empty() {
                return Ok(Some(Value::Double(0.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => {
                    // Parse leading numeric portion
                    let trimmed = s.trim();
                    let mut num_str = String::new();
                    let mut has_decimal = false;
                    let mut has_sign = false;
                    
                    for c in trimmed.chars() {
                        if c == '-' || c == '+' {
                            if has_sign || !num_str.is_empty() {
                                break;
                            }
                            has_sign = true;
                            num_str.push(c);
                        } else if c == '.' {
                            if has_decimal {
                                break;
                            }
                            has_decimal = true;
                            num_str.push(c);
                        } else if c.is_ascii_digit() {
                            num_str.push(c);
                        } else if c.is_whitespace() && num_str.is_empty() {
                            continue;
                        } else {
                            break;
                        }
                    }
                    
                    let result = num_str.parse::<f64>().unwrap_or(0.0);
                    Ok(Some(Value::Double(result)))
                }
                Value::Integer(i) => Ok(Some(Value::Double(i as f64))),
                Value::Double(d) => Ok(Some(Value::Double(d))),
                _ => Ok(Some(Value::Double(0.0)))
            }
        }

        _ => Ok(None)
    }
}
