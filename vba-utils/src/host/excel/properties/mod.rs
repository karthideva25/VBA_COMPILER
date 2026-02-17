// src/host/excel/properties/mod.rs
// Centralized property handlers for Excel objects

pub mod range_properties;
pub mod worksheet_properties;
pub mod autofilter_properties;
pub mod application;

use anyhow::Result;
use crate::context::{Context, Value};

/// Get property from any Excel object by name
pub fn get_property(
    object_type: &str,
    object_data: &str, // e.g., "A1" for Range
    property: &str,
    ctx: &mut Context,
) -> Result<Value> {
    match object_type.to_lowercase().as_str() {
        "range" => range_properties::get_range_property(object_data, property),
        "worksheet" => worksheet_properties::get_worksheet_property(object_data, property),
        "workbook" => Err(anyhow::anyhow!("Workbook properties not yet implemented")),
        "application" => application::get_property(property, ctx),
        "autofilter" => autofilter_properties::get_autofilter_property(object_data, property),
        _ => Err(anyhow::anyhow!("Unknown object type: {}", object_type)),
    }
}

/// Set property on any Excel object by name
pub fn set_property(
    object_type: &str,
    object_data: &str,
    property: &str,
    value: Value,
    ctx: &mut Context,
) -> Result<()> {
    match object_type.to_lowercase().as_str() {
        "range" => range_properties::set_range_property(object_data, property, value),
        "worksheet" => worksheet_properties::set_worksheet_property(object_data, property, value),
        "workbook" => Err(anyhow::anyhow!("Workbook properties not yet implemented")),
        "application" => application::set_property(property, value, ctx),
        "autofilter" => autofilter_properties::set_autofilter_property(object_data, property, value),
        _ => Err(anyhow::anyhow!("Unknown object type: {}", object_type)),
    }
}
