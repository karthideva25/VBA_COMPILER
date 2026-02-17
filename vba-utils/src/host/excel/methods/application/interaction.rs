// User interaction methods (dialogs, input, etc.)

use anyhow::Result;
use crate::context::Value;

pub fn call_method(method: &str, _args: &[Value]) -> Result<Value> {
    match method.to_lowercase().as_str() {
        "quit" => {
            eprintln!("🚪 Application.Quit() - closing Excel");
            Ok(Value::Empty)
        }
        "wait" => {
            eprintln!("⏱️ Application.Wait() - pausing execution");
            Ok(Value::Empty)
        }
        "inputbox" => {
            eprintln!("📝 Application.InputBox() - showing input dialog");
            Ok(Value::String(String::new()))
        }
        "ontime" => {
            eprintln!("⏰ Application.OnTime() - scheduling procedure");
            Ok(Value::Empty)
        }
        _ => Err(anyhow::anyhow!("Unknown interaction method: {}", method)),
    }
}
