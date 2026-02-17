// src/host/excel/properties/worksheet_properties.rs
// Property handlers for Worksheet object

use anyhow::{Result, bail};
use crate::context::Value;

/// Get Worksheet property by name
pub fn get_worksheet_property(data: &str, property: &str) -> Result<Value> {
    // If data is empty, get the active sheet from engine
    let name = if data.is_empty() {
        crate::host::excel::engine::get_active_sheet()
    } else {
        // Parse data format: "name:workbook_id:index" or "name::" (from Worksheets())
        let parts: Vec<&str> = data.split(':').collect();
        if parts.is_empty() {
            bail!("Invalid Worksheet data format");
        }
        parts[0].to_string()
    };
    
    let index: i32 = 1; // TODO: get from engine if available
    
    match property.to_lowercase().as_str() {
        "name" => Ok(Value::String(name.to_string())),
        "index" => Ok(Value::Integer(index as i64)),
        "visible" => Ok(Value::Boolean(true)), // TODO: get from engine
        "standardheight" => Ok(Value::Double(15.0)), // Default Excel row height
        "standardwidth" => Ok(Value::Double(8.43)), // Default Excel column width
        "autofilter" => {
            // Worksheet.AutoFilter returns the AutoFilter object for this sheet
            // Return a reference to the AutoFilter object (as an Object value)
            Ok(Value::Object(Some(Box::new(Value::String("AutoFilter".into())))))
        }
        _ => bail!("Unknown Worksheet property: {}", property),
    }
}

/// Set Worksheet property by name
pub fn set_worksheet_property(_data: &str, property: &str, _value: Value) -> Result<()> {
    match property.to_lowercase().as_str() {
        "name" => {
            eprintln!("Setting Worksheet.Name not yet implemented");
            Ok(())
        }
        "visible" => {
            eprintln!("Setting Worksheet.Visible not yet implemented");
            Ok(())
        }
        _ => bail!("Cannot set Worksheet property: {}", property),
    }
}
