// Event handler properties (OnCalculate, OnData, etc.)

use anyhow::Result;
use crate::context::Value;

pub fn get_property(property: &str) -> Result<Value> {
    match property.to_lowercase().as_str() {
        "oncalculate" | "ondata" | "ondoubleclick" | "onentry" 
        | "onsheetactivate" | "onsheetdeactivate" => {
            Ok(Value::String(String::new()))
        }
        _ => Err(anyhow::anyhow!("Unknown event property: {}", property)),
    }
}

pub fn set_property(property: &str, value: Value) -> Result<()> {
    match property.to_lowercase().as_str() {
        "oncalculate" | "ondata" | "ondoubleclick" | "onentry" 
        | "onsheetactivate" | "onsheetdeactivate" => {
            match value {
                Value::String(s) => {
                    eprintln!("📌 Application.{} = {}", property, s);
                    Ok(())
                }
                _ => Err(anyhow::anyhow!("Event property must be String")),
            }
        }
        _ => Err(anyhow::anyhow!("Cannot set event property: {}", property)),
    }
}
