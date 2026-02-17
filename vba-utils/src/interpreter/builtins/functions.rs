//! VBA Builtin Functions - Main Dispatcher
//!
//! This module acts as the entry point for handling VBA builtin function calls.
//! It delegates to category-specific modules for the actual implementation.
//!
//! Categories:
//! - strings: String manipulation functions (Len, Mid, UCase, Format, etc.)
//! - datetime: Date and time functions (Now, Date, Year, Month, etc.)
//! - math: Mathematical functions (Abs, Sqr, Round, Sin, Cos, etc.)
//! - conversion: Type conversion functions (CStr, CInt, CDbl, etc.)
//! - information: Type checking functions (VarType, TypeName, IsEmpty, etc.)
//! - interaction: User interaction and control flow (IIf, MsgBox, etc.)

use anyhow::Result;
use crate::ast::Expression;
use crate::context::{Context, Value};
use crate::interpreter::evaluate_expression;

// Import category-specific handlers
use super::strings;
use super::datetime;
use super::math;
use super::conversion;
use super::information;
use super::interaction;
use super::financial;
use super::common::value_to_string;

/// Return Ok(Some(Value)) if handled; Ok(None) to let caller try user-defined subs/funcs.
pub(crate) fn handle_builtin_call(function: &str, args: &[Expression], ctx: &mut Context)
    -> Result<Option<Value>>
{
    let func_lower = function.to_ascii_lowercase();
    
    // Try each category module in order
    // Each module returns Ok(Some(value)) if it handled the function,
    // or Ok(None) if it's not a function in that category
    
    // String functions
    if let Some(result) = strings::handle_string_function(&func_lower, args, ctx)? {
        return Ok(Some(result));
    }
    
    // Date/time functions
    if let Some(result) = datetime::handle_datetime_function(&func_lower, args, ctx)? {
        return Ok(Some(result));
    }
    
    // Math functions
    if let Some(result) = math::handle_math_function(&func_lower, args, ctx)? {
        return Ok(Some(result));
    }
    
    // Type conversion functions
    if let Some(result) = conversion::handle_conversion_function(&func_lower, args, ctx)? {
        return Ok(Some(result));
    }
    
    // Information/type checking functions
    if let Some(result) = information::handle_information_function(&func_lower, args, ctx)? {
        return Ok(Some(result));
    }
    
    // Interaction and control flow functions
    if let Some(result) = interaction::handle_interaction_function(&func_lower, args, ctx)? {
        return Ok(Some(result));
    }
    
    // Financial functions (depreciation, PV/FV, payments, IRR, etc.)
    if let Some(result) = financial::handle_financial_function(&func_lower, args, ctx)? {
        return Ok(Some(result));
    }

    // Legacy handlers for functions not yet migrated to modules
    match func_lower.as_str() {
        // MSGBOX — allow statement-style and call-style (legacy with logging)
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

        _ => Ok(None), // Not a builtin, let caller handle it
    }
}

/// Convenience wrapper to match your original `-> bool` usage
pub(crate) fn handle_builtin_call_bool(function: &str, args: &[Expression], ctx: &mut Context) -> bool {
    matches!(handle_builtin_call(function, args, ctx), Ok(Some(_)))
}

/// Set a property on a Range object
pub fn set_range_property(address: &str, property: &str, value: Value, _ctx: &mut Context) -> Result<()> {
    use crate::host::excel::engine;
    
    match property.to_lowercase().as_str() {
        "value" | "value2" => {
            // Convert the value to a string and set it in the cell
            let value_str = match value {
                Value::String(s) => s,
                Value::Integer(i) => i.to_string(),
                Value::Long(l) => l.to_string(),
                Value::LongLong(ll) => ll.to_string(),
                Value::Byte(b) => b.to_string(),
                Value::Single(s) => s.to_string(),
                Value::Double(d) => d.to_string(),
                Value::Decimal(d) => d.to_string(),
                Value::Currency(c) => c.to_string(),
                Value::Boolean(b) => if b { "TRUE" } else { "FALSE" }.to_string(),
                Value::Empty => String::new(),
                Value::Date(d) => d.to_string(),
                _ => value.as_string(),
            };
            
            // Try to set in engine, it will use in-memory storage if not initialized
            engine::set_cell_value(address, &value_str)
                .map_err(|e| anyhow::anyhow!("Failed to set cell: {}", e))?;
            Ok(())
        }
        _ => Err(anyhow::anyhow!("Unsupported Range property: {}", property)),
    }
}
