// src/host/excel/methods/autofilter_methods.rs
// Method handlers for AutoFilter object
// AutoFilter is accessed via Worksheet.AutoFilter property

use anyhow::Result;
use crate::context::Value;

/// Call method on AutoFilter object
/// Data format: "worksheet_name:workbook_id"
pub fn call_autofilter_method(_data: &str, method: &str, _args: &[Value]) -> Result<Value> {
    match method.to_lowercase().as_str() {
        "applyfilter" => {
            eprintln!("Applying AutoFilter with criteria - NOT YET IMPLEMENTED");
            Ok(Value::Empty)
        }
        "resetfilter" => {
            eprintln!("Resetting AutoFilter");
            Ok(Value::Empty)
        }
        "delete" => {
            eprintln!("Deleting AutoFilter");
            Ok(Value::Empty)
        }
        "showalldata" => {
            eprintln!("Showing all data (removing filter)");
            Ok(Value::Empty)
        }
        _ => Err(anyhow::anyhow!("Unknown AutoFilter method: {}", method)),
    }
}
