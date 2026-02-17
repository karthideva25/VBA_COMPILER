// Calculation and formula processing properties

use anyhow::Result;
use crate::context::Value;

pub fn get_property(property: &str) -> Result<Value> {
    match property.to_lowercase().as_str() {
        "calculation" => {
            // Values: Automatic, Manual, SemiAutomatic
            Ok(Value::String("Automatic".to_string()))
        }
        _ => Err(anyhow::anyhow!("Unknown calculation property: {}", property)),
    }
}

pub fn set_property(property: &str, value: Value) -> Result<()> {
    match property.to_lowercase().as_str() {
        "calculation" => {
            match value {
                Value::String(s) => {
                    eprintln!("⚙️ Application.Calculation = {}", s);
                    match s.to_lowercase().as_str() {
                        "automatic" | "manual" | "semiautomatic" => Ok(()),
                        _ => Err(anyhow::anyhow!("Invalid calculation mode: {}", s)),
                    }
                }
                _ => Err(anyhow::anyhow!("Calculation must be String")),
            }
        }
        _ => Err(anyhow::anyhow!("Cannot set calculation property: {}", property)),
    }
}
