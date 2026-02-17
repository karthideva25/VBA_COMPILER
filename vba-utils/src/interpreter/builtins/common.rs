//! Common utility functions shared across builtin function modules
//!
//! This module contains helper functions used by multiple builtin function categories.
//! 
//! ## Optional Parameter Handling Pattern
//! 
//! VBA supports optional parameters that can be omitted. Since our parser doesn't currently
//! support named arguments (param:=value) or skipping middle parameters, we handle optional
//! parameters at the interpreter level using these patterns:
//! 
//! 1. Check `args.len()` to determine which optional params were provided
//! 2. Use `get_optional_*` helpers to safely get values with defaults
//! 3. Functions with optional params in the middle should document the valid call patterns
//!
//! Example usage:
//! ```ignore
//! // For a function like InStr([start,] string1, string2, [compare])
//! let (start, str1, str2, compare) = if args.len() == 2 {
//!     (1, get_string(&args[0], ctx)?, get_string(&args[1], ctx)?, 0)
//! } else if args.len() == 3 {
//!     (get_int(&args[0], ctx)?, get_string(&args[1], ctx)?, get_string(&args[2], ctx)?, 0)
//! } else {
//!     (get_int(&args[0], ctx)?, get_string(&args[1], ctx)?, get_string(&args[2], ctx)?, get_int(&args[3], ctx)?)
//! };
//! ```

use crate::context::Value;
use crate::ast::Expression;
use crate::context::Context;
use crate::interpreter::evaluate_expression;
use anyhow::Result;

/// Convert a Value to its string representation for VBA string operations
pub(crate) fn value_to_string(val: &Value) -> String {
    match val {
        Value::String(s) => s.clone(),
        Value::Integer(n) => n.to_string(),
        Value::Long(n) => n.to_string(),
        Value::LongLong(n) => n.to_string(),
        Value::Double(n) => {
            // VBA formats doubles without unnecessary decimal places
            if n.fract() == 0.0 {
                format!("{}", *n as i64)
            } else {
                format!("{}", n)
            }
        }
        Value::Single(n) => {
            if n.fract() == 0.0 {
                format!("{}", *n as i64)
            } else {
                format!("{}", n)
            }
        }
        Value::Boolean(b) => if *b { "True".to_string() } else { "False".to_string() },
        Value::Date(d) => d.format("%m/%d/%Y").to_string(),
        Value::DateTime(dt) => dt.format("%m/%d/%Y %H:%M:%S").to_string(),
        Value::Time(t) => t.format("%H:%M:%S").to_string(),
        Value::Currency(n) => format!("{:.4}", n),
        Value::Decimal(n) => n.to_string(),
        Value::Byte(n) => n.to_string(),
        Value::Empty => String::new(),
        Value::Null => "Null".to_string(),
        Value::Object(_) => "Object".to_string(),
        Value::UserType { type_name, .. } => format!("<{} instance>", type_name),
        Value::Error(e) => format!("Error {}", e),
    }
}

/// Convert a Value to a boolean for conditional evaluation
pub(crate) fn value_to_bool(val: &Value) -> bool {
    match val {
        Value::Boolean(b) => *b,
        Value::Integer(n) => *n != 0,
        Value::Long(n) => *n != 0,
        Value::LongLong(n) => *n != 0,
        Value::Double(n) => *n != 0.0,
        Value::Single(n) => *n != 0.0,
        Value::String(s) => !s.is_empty(),
        Value::Empty | Value::Null => false,
        _ => true
    }
}

/// Convert a Value to an f64 for mathematical operations
pub(crate) fn value_to_f64(val: &Value) -> Option<f64> {
    match val {
        Value::Integer(n) => Some(*n as f64),
        Value::Long(n) => Some(*n as f64),
        Value::LongLong(n) => Some(*n as f64),
        Value::Double(n) => Some(*n),
        Value::Single(n) => Some(*n as f64),
        Value::Currency(n) => Some(*n),
        Value::Decimal(n) => Some(*n),
        Value::Byte(n) => Some(*n as f64),
        Value::Boolean(b) => Some(if *b { -1.0 } else { 0.0 }),
        Value::Empty => Some(0.0),
        Value::String(s) => s.trim().parse::<f64>().ok(),
        _ => None
    }
}

/// Convert a Value to an i64 for integer operations
pub(crate) fn value_to_i64(val: &Value) -> Option<i64> {
    match val {
        Value::Integer(n) => Some(*n),
        Value::Long(n) => Some(*n as i64),
        Value::LongLong(n) => Some(*n),
        Value::Double(n) => Some(*n as i64),
        Value::Single(n) => Some(*n as i64),
        Value::Currency(n) => Some(*n as i64),
        Value::Decimal(n) => Some(*n as i64),
        Value::Byte(n) => Some(*n as i64),
        Value::Boolean(b) => Some(if *b { -1 } else { 0 }),
        Value::Empty => Some(0),
        Value::String(s) => s.trim().parse::<i64>().ok(),
        _ => None
    }
}

// ============================================================
// OPTIONAL PARAMETER HELPER FUNCTIONS
// ============================================================

/// Get a required argument as a string value
/// Returns error if index is out of bounds or value cannot be converted to string
pub(crate) fn get_required_string(args: &[Expression], index: usize, ctx: &mut Context) -> Result<String> {
    if index >= args.len() {
        anyhow::bail!("Required argument at index {} is missing", index);
    }
    let val = evaluate_expression(&args[index], ctx)?;
    Ok(value_to_string(&val))
}

/// Get a required argument as an integer value
/// Returns error if index is out of bounds or value cannot be converted
pub(crate) fn get_required_int(args: &[Expression], index: usize, ctx: &mut Context) -> Result<i64> {
    if index >= args.len() {
        anyhow::bail!("Required argument at index {} is missing", index);
    }
    let val = evaluate_expression(&args[index], ctx)?;
    value_to_i64(&val).ok_or_else(|| anyhow::anyhow!("Cannot convert argument {} to integer", index))
}

/// Get a required argument as a floating point value  
/// Returns error if index is out of bounds or value cannot be converted
pub(crate) fn get_required_float(args: &[Expression], index: usize, ctx: &mut Context) -> Result<f64> {
    if index >= args.len() {
        anyhow::bail!("Required argument at index {} is missing", index);
    }
    let val = evaluate_expression(&args[index], ctx)?;
    value_to_f64(&val).ok_or_else(|| anyhow::anyhow!("Cannot convert argument {} to float", index))
}

/// Get an optional argument as a string, with a default value
pub(crate) fn get_optional_string(args: &[Expression], index: usize, default: &str, ctx: &mut Context) -> Result<String> {
    if index >= args.len() {
        return Ok(default.to_string());
    }
    let val = evaluate_expression(&args[index], ctx)?;
    Ok(value_to_string(&val))
}

/// Get an optional argument as an integer, with a default value
pub(crate) fn get_optional_int(args: &[Expression], index: usize, default: i64, ctx: &mut Context) -> Result<i64> {
    if index >= args.len() {
        return Ok(default);
    }
    let val = evaluate_expression(&args[index], ctx)?;
    Ok(value_to_i64(&val).unwrap_or(default))
}

/// Get an optional argument as a float, with a default value
pub(crate) fn get_optional_float(args: &[Expression], index: usize, default: f64, ctx: &mut Context) -> Result<f64> {
    if index >= args.len() {
        return Ok(default);
    }
    let val = evaluate_expression(&args[index], ctx)?;
    Ok(value_to_f64(&val).unwrap_or(default))
}

/// Get an optional argument as a boolean, with a default value
pub(crate) fn get_optional_bool(args: &[Expression], index: usize, default: bool, ctx: &mut Context) -> Result<bool> {
    if index >= args.len() {
        return Ok(default);
    }
    let val = evaluate_expression(&args[index], ctx)?;
    Ok(value_to_bool(&val))
}

/// Get an argument as a raw Value (useful when you need the actual type)
pub(crate) fn get_arg_value(args: &[Expression], index: usize, ctx: &mut Context) -> Result<Option<Value>> {
    if index >= args.len() {
        return Ok(None);
    }
    Ok(Some(evaluate_expression(&args[index], ctx)?))
}
