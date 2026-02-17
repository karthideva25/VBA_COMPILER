// src/host/excel/methods/worksheet_methods.rs
// Method handlers for Worksheet object

use anyhow::Result;
use crate::context::Value;

/// Call method on Worksheet object
pub fn call_worksheet_method(data: &str, method: &str, _args: &[Value]) -> Result<Value> {
    let parts: Vec<&str> = data.split(':').collect();
    let name = parts.get(0).copied().unwrap_or("Sheet");
    
    match method.to_lowercase().as_str() {
        "activate" => {
            eprintln!("Activated worksheet: {}", name);
            Ok(Value::Empty)
        }
        "delete" => {
            eprintln!("Deleting worksheet: {} - NOT YET IMPLEMENTED", name);
            Ok(Value::Empty)
        }
        "copy" => {
            eprintln!("Copying worksheet: {} - NOT YET IMPLEMENTED", name);
            Ok(Value::Empty)
        }
        "move" => {
            eprintln!("Moving worksheet: {} - NOT YET IMPLEMENTED", name);
            Ok(Value::Empty)
        }
        "select" => {
            eprintln!("Selecting worksheet: {}", name);
            Ok(Value::Empty)
        }
        _ => Err(anyhow::anyhow!("Unknown Worksheet method: {}", method)),
    }
}
