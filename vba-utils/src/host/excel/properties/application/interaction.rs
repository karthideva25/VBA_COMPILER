// Display, alerts, and event handling properties

use anyhow::Result;
use crate::context::Value;

pub fn get_property(property: &str) -> Result<Value> {
    match property.to_lowercase().as_str() {
        "displayalerts" => {
            // Get from engine if available
            Ok(Value::Boolean(true))
        }
        "screenupdating" => {
            Ok(Value::Boolean(true))
        }
        "enableevents" => {
            Ok(Value::Boolean(true))
        }
        _ => Err(anyhow::anyhow!("Unknown interaction property: {}", property)),
    }
}

pub fn set_property(property: &str, value: Value) -> Result<()> {
    match property.to_lowercase().as_str() {
        "displayalerts" => {
            match value {
                Value::Boolean(b) => {
                    eprintln!("📢 Application.DisplayAlerts = {}", b);
                    Ok(())
                }
                _ => Err(anyhow::anyhow!("DisplayAlerts must be Boolean")),
            }
        }
        "screenupdating" => {
            match value {
                Value::Boolean(b) => {
                    eprintln!("📊 Application.ScreenUpdating = {}", b);
                    Ok(())
                }
                _ => Err(anyhow::anyhow!("ScreenUpdating must be Boolean")),
            }
        }
        "enableevents" => {
            match value {
                Value::Boolean(b) => {
                    eprintln!("⚡ Application.EnableEvents = {}", b);
                    Ok(())
                }
                _ => Err(anyhow::anyhow!("EnableEvents must be Boolean")),
            }
        }
        _ => Err(anyhow::anyhow!("Cannot set interaction property: {}", property)),
    }
}
