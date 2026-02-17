//! VBA Err Object Functions
//! 
//! This module contains all VBA Err object related functions including:
//! - Err() - Returns the Err object
//! - Err.Number - Error number (get/set)
//! - Err.Description - Error description (get/set)
//! - Err.Source - Error source (get/set)
//! - Err.Clear - Clears all error properties
//! - Err.Raise - Generates a runtime error

use anyhow::Result;
use crate::ast::Expression;
use crate::context::{Context, Value, ErrObject};
use crate::interpreter::evaluate_expression;

/// Handle Err() function call (returns default property = Number)
pub(crate) fn handle_err_function(args: &[Expression], ctx: &mut Context) -> Result<Option<Value>> {
    // When called as Err() without property access, return the error number
    // This is the default property of Err object
    let _ = args; // Err() takes no arguments
    let number = ctx.err.as_ref().map(|e| e.number).unwrap_or(0);
    Ok(Some(Value::Integer(number.into())))
}

/// Handle Err object method calls (Raise, Clear, Number, Description, Source)
pub(crate) fn handle_err_method(method: &str, args: &[Expression], ctx: &mut Context) -> Result<Option<Value>> {
    match method.to_ascii_lowercase().as_str() {
        // ERR.RAISE — Generates a runtime error
        // Syntax: Err.Raise Number, [Source], [Description], [HelpFile], [HelpContext]
        "raise" => {
            if args.is_empty() {
                anyhow::bail!("Err.Raise requires at least a Number argument");
            }
            
            // Required: Number
            let number = value_to_i32(&evaluate_expression(&args[0], ctx)?);
            
            // Optional: Source (default: current project name)
            let source = if args.len() > 1 {
                evaluate_expression(&args[1], ctx)?.as_string()
            } else {
                "VBA".to_string()
            };
            
            // Optional: Description (default: error message for number)
            let description = if args.len() > 2 {
                evaluate_expression(&args[2], ctx)?.as_string()
            } else {
                get_default_error_description(number)
            };
            
            // HelpFile and HelpContext are ignored (args[3] and args[4])
            // as they're not relevant for this implementation
            
            // Set the error in context
            ctx.err = Some(ErrObject {
                number,
                description,
                source,
            });
            
            // Return error indication - the calling code should handle this
            Ok(Some(Value::Error(number)))
        }
        
        // ERR.CLEAR — Clears all property settings of the Err object
        "clear" => {
            ctx.err = None;
            ctx.resume_valid = false;
            Ok(Some(Value::Empty))
        }
        
        // ERR.NUMBER — Get/Set error number (getter handled in PropertyAccess)
        "number" => {
            if !args.is_empty() {
                // Setter: Err.Number = value
                let number = value_to_i32(&evaluate_expression(&args[0], ctx)?);
                if let Some(ref mut err) = ctx.err {
                    err.number = number;
                } else {
                    ctx.err = Some(ErrObject {
                        number,
                        description: String::new(),
                        source: String::new(),
                    });
                }
            }
            let n = ctx.err.as_ref().map(|e| e.number).unwrap_or(0);
            Ok(Some(Value::Integer(n.into())))
        }
        
        // ERR.DESCRIPTION — Get/Set error description
        "description" => {
            if !args.is_empty() {
                // Setter
                let desc = evaluate_expression(&args[0], ctx)?.as_string();
                if let Some(ref mut err) = ctx.err {
                    err.description = desc;
                } else {
                    ctx.err = Some(ErrObject {
                        number: 0,
                        description: desc,
                        source: String::new(),
                    });
                }
            }
            let d = ctx.err.as_ref().map(|e| e.description.clone()).unwrap_or_default();
            Ok(Some(Value::String(d)))
        }
        
        // ERR.SOURCE — Get/Set error source
        "source" => {
            if !args.is_empty() {
                // Setter
                let src = evaluate_expression(&args[0], ctx)?.as_string();
                if let Some(ref mut err) = ctx.err {
                    err.source = src;
                } else {
                    ctx.err = Some(ErrObject {
                        number: 0,
                        description: String::new(),
                        source: src,
                    });
                }
            }
            let s = ctx.err.as_ref().map(|e| e.source.clone()).unwrap_or_default();
            Ok(Some(Value::String(s)))
        }
        
        _ => Ok(None)
    }
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================

fn value_to_i32(val: &Value) -> i32 {
    match val {
        Value::Integer(i) => *i as i32,
        Value::Long(l) => *l,
        Value::LongLong(ll) => *ll as i32,
        Value::Double(d) => *d as i32,
        Value::Single(s) => *s as i32,
        Value::Currency(c) => *c as i32,
        Value::String(s) => s.parse::<i32>().unwrap_or(0),
        Value::Boolean(b) => if *b { -1 } else { 0 },
        Value::Byte(b) => *b as i32,
        _ => 0
    }
}

/// Get default error description for common VBA error numbers
pub fn get_default_error_description(err_num: i32) -> String {
    match err_num {
        0 => String::new(),
        3 => "Return without GoSub".to_string(),
        5 => "Invalid procedure call or argument".to_string(),
        6 => "Overflow".to_string(),
        7 => "Out of memory".to_string(),
        9 => "Subscript out of range".to_string(),
        10 => "This array is fixed or temporarily locked".to_string(),
        11 => "Division by zero".to_string(),
        13 => "Type mismatch".to_string(),
        14 => "Out of string space".to_string(),
        17 => "Can't perform requested operation".to_string(),
        18 => "User interrupt occurred".to_string(),
        20 => "Resume without error".to_string(),
        28 => "Out of stack space".to_string(),
        35 => "Sub or Function not defined".to_string(),
        48 => "Error in loading DLL".to_string(),
        51 => "Internal error".to_string(),
        52 => "Bad file name or number".to_string(),
        53 => "File not found".to_string(),
        54 => "Bad file mode".to_string(),
        55 => "File already open".to_string(),
        57 => "Device I/O error".to_string(),
        58 => "File already exists".to_string(),
        61 => "Disk full".to_string(),
        62 => "Input past end of file".to_string(),
        68 => "Device unavailable".to_string(),
        70 => "Permission denied".to_string(),
        71 => "Disk not ready".to_string(),
        75 => "Path/File access error".to_string(),
        76 => "Path not found".to_string(),
        91 => "Object variable or With block variable not set".to_string(),
        92 => "For loop not initialized".to_string(),
        94 => "Invalid use of Null".to_string(),
        424 => "Object required".to_string(),
        429 => "ActiveX component can't create object".to_string(),
        438 => "Object doesn't support this property or method".to_string(),
        440 => "Automation error".to_string(),
        449 => "Argument not optional".to_string(),
        450 => "Wrong number of arguments or invalid property assignment".to_string(),
        1004 => "Application-defined or object-defined error".to_string(),
        _ => format!("Application-defined or object-defined error {}", err_num)
    }
}
