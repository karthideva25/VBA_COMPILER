// src/host/excel/methods/mod.rs
// Centralized method handlers for Excel objects

pub mod range_methods;
pub mod worksheet_methods;
pub mod autofilter_methods;
pub mod application;

use anyhow::Result;
use crate::context::Value;

/// Call method on any Excel object
pub fn call_method(
    object_type: &str,
    object_data: &str, // e.g., "A1" for Range
    method: &str,
    args: &[Value],
) -> Result<Value> {
    match object_type.to_lowercase().as_str() {
        "range" => range_methods::call_range_method(object_data, method, args),
        "worksheet" => worksheet_methods::call_worksheet_method(object_data, method, args),
        "workbook" => Err(anyhow::anyhow!("Workbook methods not yet implemented")),
        "application" => application::call_method(method, args, &mut crate::context::Context::default()),
        "autofilter" => autofilter_methods::call_autofilter_method(object_data, method, args),
        _ => Err(anyhow::anyhow!("Unknown object type: {}", object_type)),
    }
}
