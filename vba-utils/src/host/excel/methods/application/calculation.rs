// Calculation and formula evaluation methods

use anyhow::Result;
use crate::context::Value;

pub fn call_method(method: &str, _args: &[Value]) -> Result<Value> {
    match method.to_lowercase().as_str() {
        "calculate" => {
            eprintln!("🧮 Application.Calculate() - recalculating all open workbooks");
            Ok(Value::Empty)
        }
        "calculatefull" => {
            eprintln!("🧮 Application.CalculateFull() - full recalculation (forces rebuild)");
            Ok(Value::Empty)
        }
        _ => Err(anyhow::anyhow!("Unknown calculation method: {}", method)),
    }
}
