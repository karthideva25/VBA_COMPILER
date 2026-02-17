// User and creator information properties

use anyhow::Result;
use crate::context::Value;

pub fn get_property(property: &str) -> Result<Value> {
    match property.to_lowercase().as_str() {
        "username" => Ok(Value::String("User".to_string())),
        "useremailid" => Ok(Value::String(String::new())),
        "creatorname" => Ok(Value::String(String::new())),
        "creatoremailid" => Ok(Value::String(String::new())),
        _ => Err(anyhow::anyhow!("Unknown metadata property: {}", property)),
    }
}

pub fn set_property(property: &str, value: Value) -> Result<()> {
    match property.to_lowercase().as_str() {
        "username" | "useremailid" | "creatorname" | "creatoremailid" => {
            match value {
                Value::String(s) => {
                    eprintln!("👤 Application.{} = {}", property, s);
                    Ok(())
                }
                _ => Err(anyhow::anyhow!("Metadata property must be String")),
            }
        }
        _ => Err(anyhow::anyhow!("Cannot set metadata property: {}", property)),
    }
}
