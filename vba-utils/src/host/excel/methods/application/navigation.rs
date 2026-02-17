// Navigation and selection methods

use anyhow::Result;
use crate::context::Value;

pub fn call_method(method: &str, _args: &[Value]) -> Result<Value> {
    match method.to_lowercase().as_str() {
        "goto" => {
            eprintln!("🔍 Application.Goto() - navigate to cell/range");
            Ok(Value::Empty)
        }
        "activateprevious" => {
            eprintln!("⬅️ Application.ActivatePrevious() - activate previous workbook");
            Ok(Value::Empty)
        }
        "activatenext" => {
            eprintln!("➡️ Application.ActivateNext() - activate next workbook");
            Ok(Value::Empty)
        }
        _ => Err(anyhow::anyhow!("Unknown navigation method: {}", method)),
    }
}
