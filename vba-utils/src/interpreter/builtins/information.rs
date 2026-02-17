//! VBA Information Functions
//! 
//! This module contains all VBA information/type checking functions including:
//! - IsArray, IsDate, IsEmpty, IsError, IsMissing, IsNull, IsNumeric, IsObject
//! - VarType, TypeName
//! - QBColor, RGB

use anyhow::Result;
use chrono::NaiveDate;
use crate::ast::Expression;
use crate::context::{Context, Value};
use crate::interpreter::evaluate_expression;

/// Handle information-related builtin function calls
pub(crate) fn handle_information_function(function: &str, args: &[Expression], ctx: &mut Context) -> Result<Option<Value>> {
    match function {
        // ============================================================
        // TYPE CHECKING FUNCTIONS
        // ============================================================

        // VARTYPE — Returns an Integer indicating the subtype of a variable
        "vartype" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0))); // vbEmpty
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let type_code = val.get_var_type() as i64;
            Ok(Some(Value::Integer(type_code)))
        }

        // TYPENAME — Returns a String indicating the subtype of a variable
        "typename" => {
            if args.is_empty() {
                return Ok(Some(Value::String("Empty".to_string())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            Ok(Some(Value::String(val.type_name().to_string())))
        }

        // ISEMPTY — Returns True if variable is uninitialized (Empty)
        "isempty" => {
            if args.is_empty() {
                return Ok(Some(Value::Boolean(true)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            Ok(Some(Value::Boolean(val.is_empty())))
        }

        // ISNULL — Returns True if expression is Null
        "isnull" => {
            if args.is_empty() {
                return Ok(Some(Value::Boolean(false)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            Ok(Some(Value::Boolean(val.is_null())))
        }

        // ISNUMERIC — Returns True if expression can be evaluated as a number
        "isnumeric" => {
            if args.is_empty() {
                return Ok(Some(Value::Boolean(false)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let is_numeric = match val {
                Value::Integer(_) | Value::Long(_) | Value::LongLong(_) |
                Value::Double(_) | Value::Single(_) | Value::Currency(_) |
                Value::Decimal(_) | Value::Byte(_) => true,
                Value::Boolean(_) => true, // Booleans are numeric in VBA
                Value::String(s) => s.trim().parse::<f64>().is_ok(),
                Value::Empty => true, // Empty converts to 0
                _ => false
            };
            Ok(Some(Value::Boolean(is_numeric)))
        }

        // ISDATE — Returns True if expression can be converted to a date
        "isdate" => {
            if args.is_empty() {
                return Ok(Some(Value::Boolean(false)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let is_date = match val {
                Value::Date(_) | Value::DateTime(_) => true,
                Value::String(s) => {
                    let formats = ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d"];
                    formats.iter().any(|fmt| NaiveDate::parse_from_str(&s, fmt).is_ok())
                }
                _ => false
            };
            Ok(Some(Value::Boolean(is_date)))
        }

        // ISOBJECT — Returns True if expression is an Object
        "isobject" => {
            if args.is_empty() {
                return Ok(Some(Value::Boolean(false)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let is_object = matches!(val, Value::Object(_));
            Ok(Some(Value::Boolean(is_object)))
        }

        // ISARRAY — Returns True if variable is an array
        "isarray" => {
            // For now, arrays are not fully supported
            // TODO: When arrays are implemented, check if value is array type
            Ok(Some(Value::Boolean(false)))
        }

        // ISMISSING — Returns True if optional argument was not passed
        "ismissing" => {
            if args.is_empty() {
                return Ok(Some(Value::Boolean(true)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            // If argument was passed, it's not missing (Empty indicates missing optional)
            Ok(Some(Value::Boolean(matches!(val, Value::Empty))))
        }

        // ISERROR — Returns True if expression is an error value
        "iserror" => {
            if args.is_empty() {
                return Ok(Some(Value::Boolean(false)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            // Check for Value::Error (from CVErr)
            Ok(Some(Value::Boolean(matches!(val, Value::Error(_)))))
        }

        // ISNOTHING — Returns True if object variable is Nothing
        "isnothing" => {
            if args.is_empty() {
                return Ok(Some(Value::Boolean(true)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let is_nothing = matches!(val, Value::Object(None));
            Ok(Some(Value::Boolean(is_nothing)))
        }

        // ============================================================
        // COLOR FUNCTIONS
        // ============================================================

        // QBCOLOR — Returns an RGB color value from a QBasic color number (0-15)
        "qbcolor" => {
            if args.is_empty() {
                return Ok(Some(Value::Long(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let color_num = value_to_i32(&val).clamp(0, 15);
            
            // QBasic color palette (0-15 map to specific RGB values)
            let rgb = match color_num {
                0 => 0x000000,   // Black
                1 => 0x800000,   // Blue (dark)
                2 => 0x008000,   // Green (dark)
                3 => 0x808000,   // Cyan (dark)
                4 => 0x000080,   // Red (dark)
                5 => 0x800080,   // Magenta (dark)
                6 => 0x008080,   // Yellow (dark) / Brown
                7 => 0xC0C0C0,   // White (light gray)
                8 => 0x808080,   // Gray
                9 => 0xFF0000,   // Blue (bright)
                10 => 0x00FF00,  // Green (bright)
                11 => 0xFFFF00,  // Cyan (bright)
                12 => 0x0000FF,  // Red (bright)
                13 => 0xFF00FF,  // Magenta (bright)
                14 => 0x00FFFF,  // Yellow (bright)
                15 => 0xFFFFFF,  // White (bright)
                _ => 0x000000,
            };
            Ok(Some(Value::Long(rgb)))
        }

        // RGB — Returns a Long whole number representing an RGB color value
        "rgb" => {
            let red = if !args.is_empty() {
                value_to_i32(&evaluate_expression(&args[0], ctx)?).clamp(0, 255)
            } else { 0 };
            
            let green = if args.len() > 1 {
                value_to_i32(&evaluate_expression(&args[1], ctx)?).clamp(0, 255)
            } else { 0 };
            
            let blue = if args.len() > 2 {
                value_to_i32(&evaluate_expression(&args[2], ctx)?).clamp(0, 255)
            } else { 0 };
            
            // VBA RGB format: Blue in high byte, Green in middle, Red in low
            // RGB = Red + (Green * 256) + (Blue * 65536)
            let rgb_value = red + (green * 256) + (blue * 65536);
            Ok(Some(Value::Long(rgb_value)))
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
