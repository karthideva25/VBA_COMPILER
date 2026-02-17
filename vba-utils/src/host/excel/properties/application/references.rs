// Reference style and clipboard mode properties

use anyhow::Result;
use crate::context::Value;

pub fn get_property(property: &str) -> Result<Value> {
    match property.to_lowercase().as_str() {
        "referencestyle" => {
            // 1 = A1 style, 2 = R1C1 style
            Ok(Value::Integer(1))
        }
        "cutcopymode" => {
            // 0 = none, 1 = cut, 2 = copy
            Ok(Value::Integer(0))
        }
        _ => Err(anyhow::anyhow!("Unknown reference property: {}", property)),
    }
}

pub fn set_property(property: &str, value: Value) -> Result<()> {
    match property.to_lowercase().as_str() {
        "referencestyle" => {
            match value {
                Value::Integer(i) if i == 1 || i == 2 => {
                    let style = if i == 1 { "A1" } else { "R1C1" };
                    eprintln!("📋 Application.ReferenceStyle = {} ({})", i, style);
                    Ok(())
                }
                _ => Err(anyhow::anyhow!("ReferenceStyle must be 1 or 2")),
            }
        }
        "cutcopymode" => {
            match value {
                Value::Integer(i) if i >= 0 && i <= 2 => {
                    eprintln!("✂️ Application.CutCopyMode = {}", i);
                    Ok(())
                }
                _ => Err(anyhow::anyhow!("CutCopyMode must be 0, 1, or 2")),
            }
        }
        _ => Err(anyhow::anyhow!("Cannot set reference property: {}", property)),
    }
}
