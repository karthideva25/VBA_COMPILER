// src/host/excel/properties/autofilter_properties.rs
// Property handlers for AutoFilter object
// AutoFilter is accessed via Worksheet.AutoFilter property

use anyhow::{Result, bail};
use crate::context::Value;

/// Get AutoFilter property by name
/// Data format: "worksheet_name:workbook_id"
pub fn get_autofilter_property(_data: &str, property: &str) -> Result<Value> {
    match property.to_lowercase().as_str() {
        "enabled" => Ok(Value::Boolean(true)),
        "range" => {
            // AutoFilter.Range returns the range the filter is applied to
            Ok(Value::String("A1:Z100".to_string())) // TODO: get from engine
        }
        _ => bail!("Unknown AutoFilter property: {}", property),
    }
}

/// Set AutoFilter property by name
pub fn set_autofilter_property(_data: &str, property: &str, _value: Value) -> Result<()> {
    match property.to_lowercase().as_str() {
        "enabled" => {
            eprintln!("Setting AutoFilter.Enabled not yet implemented");
            Ok(())
        }
        _ => bail!("Cannot set AutoFilter property: {}", property),
    }
}
