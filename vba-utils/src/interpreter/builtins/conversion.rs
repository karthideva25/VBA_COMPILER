//! VBA Type Conversion Functions
//! 
//! This module contains all VBA type conversion functions including:
//! - CBool, CByte, CCur, CDate, CDbl, CDec
//! - CInt, CLng, CLngLng, CSng, CStr, CVar
//! - CVDate, CVErr
//! - Error, Error$
//! - Str, Str$

use anyhow::Result;
use chrono::NaiveDate;
use crate::ast::Expression;
use crate::context::{Context, Value};
use crate::interpreter::evaluate_expression;
use super::common::value_to_string;

/// Handle type conversion builtin function calls
pub(crate) fn handle_conversion_function(function: &str, args: &[Expression], ctx: &mut Context) -> Result<Option<Value>> {
    match function {
        // ============================================================
        // STRING CONVERSION
        // ============================================================

        // CSTR — Convert to String
        "cstr" => {
            if args.is_empty() {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            Ok(Some(Value::String(value_to_string(&val))))
        }

        // ============================================================
        // INTEGER CONVERSIONS
        // ============================================================

        // CINT — Convert to Integer (rounds to nearest even)
        "cint" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = value_to_f64(&val);
            // VBA rounds to nearest even (banker's rounding)
            Ok(Some(Value::Integer(f.round() as i64)))
        }

        // CLNG — Convert to Long
        "clng" => {
            if args.is_empty() {
                return Ok(Some(Value::Long(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = value_to_f64(&val);
            Ok(Some(Value::Long(f.round() as i32)))
        }

        // CLNGLNG — Convert to LongLong (64-bit)
        "clnglng" => {
            if args.is_empty() {
                return Ok(Some(Value::LongLong(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = value_to_f64(&val);
            Ok(Some(Value::LongLong(f.round() as i64)))
        }

        // CBYTE — Convert to Byte (0-255)
        "cbyte" => {
            if args.is_empty() {
                return Ok(Some(Value::Byte(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = value_to_f64(&val);
            let byte_val = f.round().clamp(0.0, 255.0) as u8;
            Ok(Some(Value::Byte(byte_val)))
        }

        // ============================================================
        // FLOATING POINT CONVERSIONS
        // ============================================================

        // CDBL — Convert to Double
        "cdbl" => {
            if args.is_empty() {
                return Ok(Some(Value::Double(0.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = value_to_f64(&val);
            Ok(Some(Value::Double(f)))
        }

        // CSNG — Convert to Single
        "csng" => {
            if args.is_empty() {
                return Ok(Some(Value::Single(0.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = value_to_f64(&val);
            Ok(Some(Value::Single(f as f32)))
        }

        // CDEC — Convert to Decimal
        "cdec" => {
            if args.is_empty() {
                return Ok(Some(Value::Decimal(0.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = value_to_f64(&val);
            Ok(Some(Value::Decimal(f)))
        }

        // ============================================================
        // CURRENCY AND SPECIAL TYPES
        // ============================================================

        // CCUR — Convert to Currency
        "ccur" => {
            if args.is_empty() {
                return Ok(Some(Value::Currency(0.0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = value_to_f64(&val);
            // Currency has 4 decimal places
            let rounded = (f * 10000.0).round() / 10000.0;
            Ok(Some(Value::Currency(rounded)))
        }

        // CBOOL — Convert to Boolean
        "cbool" => {
            if args.is_empty() {
                return Ok(Some(Value::Boolean(false)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let b = match val {
                Value::Boolean(b) => b,
                Value::Integer(i) => i != 0,
                Value::Long(l) => l != 0,
                Value::LongLong(ll) => ll != 0,
                Value::Double(d) => d != 0.0,
                Value::Single(s) => s != 0.0,
                Value::String(s) => {
                    let lower = s.to_lowercase();
                    lower == "true" || lower == "yes" || lower == "-1"
                }
                Value::Empty => false,
                Value::Null => false,
                _ => false
            };
            Ok(Some(Value::Boolean(b)))
        }

        // CDATE — Convert to Date
        "cdate" => {
            if args.is_empty() {
                return Ok(Some(Value::Empty));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::Date(d) => Ok(Some(Value::Date(d))),
                Value::String(s) => {
                    // Try common date formats
                    let formats = [
                        "%Y-%m-%d",
                        "%m/%d/%Y",
                        "%d/%m/%Y",
                        "%Y/%m/%d",
                        "%B %d, %Y",
                    ];
                    for fmt in formats.iter() {
                        if let Ok(date) = NaiveDate::parse_from_str(&s, fmt) {
                            return Ok(Some(Value::Date(date)));
                        }
                    }
                    Ok(Some(Value::Empty))
                }
                Value::Integer(i) => {
                    // VBA serial date (days since Dec 30, 1899)
                    let base = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
                    if let Some(date) = base.checked_add_signed(chrono::Duration::days(i)) {
                        Ok(Some(Value::Date(date)))
                    } else {
                        Ok(Some(Value::Empty))
                    }
                }
                Value::Double(d) => {
                    let base = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
                    if let Some(date) = base.checked_add_signed(chrono::Duration::days(d.trunc() as i64)) {
                        Ok(Some(Value::Date(date)))
                    } else {
                        Ok(Some(Value::Empty))
                    }
                }
                _ => Ok(Some(Value::Empty))
            }
        }

        // CVAR — Convert to Variant (returns value as-is)
        "cvar" => {
            if args.is_empty() {
                return Ok(Some(Value::Empty));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            Ok(Some(val))
        }

        // ============================================================
        // VARIANT DATE AND ERROR CONVERSIONS
        // ============================================================

        // CVDATE — Convert to Variant containing Date (legacy, same as CDate)
        "cvdate" => {
            if args.is_empty() {
                return Ok(Some(Value::Empty));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::Date(d) => Ok(Some(Value::Date(d))),
                Value::DateTime(dt) => Ok(Some(Value::Date(dt.date()))),
                Value::String(s) => {
                    let formats = [
                        "%Y-%m-%d",
                        "%m/%d/%Y",
                        "%d/%m/%Y",
                        "%Y/%m/%d",
                        "%B %d, %Y",
                    ];
                    for fmt in formats.iter() {
                        if let Ok(date) = NaiveDate::parse_from_str(&s, fmt) {
                            return Ok(Some(Value::Date(date)));
                        }
                    }
                    Ok(Some(Value::Empty))
                }
                Value::Integer(i) => {
                    let base = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
                    if let Some(date) = base.checked_add_signed(chrono::Duration::days(i)) {
                        Ok(Some(Value::Date(date)))
                    } else {
                        Ok(Some(Value::Empty))
                    }
                }
                Value::Double(d) => {
                    let base = NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
                    if let Some(date) = base.checked_add_signed(chrono::Duration::days(d.trunc() as i64)) {
                        Ok(Some(Value::Date(date)))
                    } else {
                        Ok(Some(Value::Empty))
                    }
                }
                _ => Ok(Some(Value::Empty))
            }
        }

        // CVERR — Create an Error value from an error number
        "cverr" => {
            if args.is_empty() {
                return Ok(Some(Value::Error(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let err_num = match val {
                Value::Integer(i) => i as i32,
                Value::Long(l) => l,
                Value::Double(d) => d as i32,
                _ => 0
            };
            Ok(Some(Value::Error(err_num)))
        }

        // ============================================================
        // ERROR MESSAGE FUNCTIONS
        // ============================================================

        // ERROR / ERROR$ — Returns error message for error number
        "error" | "error$" => {
            if args.is_empty() {
                // Return current error message (if any)
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let err_num = match val {
                Value::Integer(i) => i as i32,
                Value::Long(l) => l,
                Value::Double(d) => d as i32,
                _ => 0
            };
            let msg = get_error_message(err_num);
            Ok(Some(Value::String(msg)))
        }

        // ============================================================
        // STRING CONVERSION
        // ============================================================

        // STR / STR$ — Convert number to string with leading space for positive
        "str" | "str$" => {
            if args.is_empty() {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let f = value_to_f64(&val);
            // VBA Str adds a leading space for positive numbers
            let result = if f >= 0.0 {
                format!(" {}", f)
            } else {
                format!("{}", f)
            };
            Ok(Some(Value::String(result)))
        }

        _ => Ok(None)
    }
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================

fn value_to_f64(val: &Value) -> f64 {
    match val {
        Value::Integer(i) => *i as f64,
        Value::Long(l) => *l as f64,
        Value::LongLong(ll) => *ll as f64,
        Value::Double(d) => *d,
        Value::Single(s) => *s as f64,
        Value::Currency(c) => *c,
        Value::String(s) => s.parse::<f64>().unwrap_or(0.0),
        Value::Boolean(b) => if *b { -1.0 } else { 0.0 },
        Value::Byte(b) => *b as f64,
        _ => 0.0
    }
}

/// Get VBA error message for error number
fn get_error_message(err_num: i32) -> String {
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
        16 => "Expression too complex".to_string(),
        17 => "Can't perform requested operation".to_string(),
        18 => "User interrupt occurred".to_string(),
        20 => "Resume without error".to_string(),
        28 => "Out of stack space".to_string(),
        35 => "Sub or Function not defined".to_string(),
        48 => "Error in loading DLL".to_string(),
        49 => "Bad DLL calling convention".to_string(),
        51 => "Internal error".to_string(),
        52 => "Bad file name or number".to_string(),
        53 => "File not found".to_string(),
        54 => "Bad file mode".to_string(),
        55 => "File already open".to_string(),
        57 => "Device I/O error".to_string(),
        58 => "File already exists".to_string(),
        59 => "Bad record length".to_string(),
        61 => "Disk full".to_string(),
        62 => "Input past end of file".to_string(),
        63 => "Bad record number".to_string(),
        67 => "Too many files".to_string(),
        68 => "Device unavailable".to_string(),
        70 => "Permission denied".to_string(),
        71 => "Disk not ready".to_string(),
        74 => "Can't rename with different drive".to_string(),
        75 => "Path/File access error".to_string(),
        76 => "Path not found".to_string(),
        91 => "Object variable or With block variable not set".to_string(),
        92 => "For loop not initialized".to_string(),
        93 => "Invalid pattern string".to_string(),
        94 => "Invalid use of Null".to_string(),
        97 => "Can't call Friend function on object which is not an instance of defining class".to_string(),
        98 => "A property or method call cannot include a reference to a private object".to_string(),
        321 => "Invalid file format".to_string(),
        322 => "Can't create necessary temporary file".to_string(),
        325 => "Invalid format in resource file".to_string(),
        380 => "Invalid property value".to_string(),
        381 => "Invalid property array index".to_string(),
        382 => "Set not supported at runtime".to_string(),
        383 => "Set not supported (read-only property)".to_string(),
        385 => "Need property array index".to_string(),
        387 => "Set not permitted".to_string(),
        393 => "Get not supported at runtime".to_string(),
        394 => "Get not supported (write-only property)".to_string(),
        422 => "Property not found".to_string(),
        423 => "Property or method not found".to_string(),
        424 => "Object required".to_string(),
        429 => "ActiveX component can't create object".to_string(),
        430 => "Class doesn't support Automation or doesn't support expected interface".to_string(),
        432 => "File name or class name not found during Automation operation".to_string(),
        438 => "Object doesn't support this property or method".to_string(),
        440 => "Automation error".to_string(),
        442 => "Connection to type library or object library for remote process has been lost".to_string(),
        443 => "Automation object does not have a default value".to_string(),
        445 => "Object doesn't support this action".to_string(),
        446 => "Object doesn't support named arguments".to_string(),
        447 => "Object doesn't support current locale setting".to_string(),
        448 => "Named argument not found".to_string(),
        449 => "Argument not optional".to_string(),
        450 => "Wrong number of arguments or invalid property assignment".to_string(),
        451 => "Property let procedure not defined and property get procedure did not return an object".to_string(),
        452 => "Invalid ordinal".to_string(),
        453 => "Specified DLL function not found".to_string(),
        454 => "Code resource not found".to_string(),
        455 => "Code resource lock error".to_string(),
        457 => "This key is already associated with an element of this collection".to_string(),
        458 => "Variable uses an Automation type not supported in Visual Basic".to_string(),
        459 => "Object or class does not support the set of events".to_string(),
        460 => "Invalid clipboard format".to_string(),
        461 => "Method or data member not found".to_string(),
        462 => "The remote server machine does not exist or is unavailable".to_string(),
        463 => "Class not registered on local machine".to_string(),
        481 => "Invalid picture".to_string(),
        482 => "Printer error".to_string(),
        735 => "Can't save file to TEMP".to_string(),
        744 => "Search text not found".to_string(),
        746 => "Replacements too long".to_string(),
        1004 => "Application-defined or object-defined error".to_string(),
        _ => format!("Unknown error {}", err_num)
    }
}
