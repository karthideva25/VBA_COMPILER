// src/host/excel/objects/mod.rs
// Unified Excel object handling and dispatch

use anyhow::Result;
use crate::context::{Context, Value};

// Active objects (used by COM registry and interpreter)
pub mod application;
pub mod range;

// Re-export key types for convenience
pub use range::{ExcelRange, RangeBuilder, indices_to_address, column_index_to_letter};

/// Unified dispatcher for Excel object properties and methods
/// Handles: Range, Worksheet, Workbook, Application, AutoFilter, etc.
pub fn dispatch_property_get(
    object_type: &str,
    object_data: &str,
    property: &str,
    ctx: &mut Context,
) -> Result<Value> {
    super::properties::get_property(object_type, object_data, property, ctx)
}

pub fn dispatch_property_set(
    object_type: &str,
    object_data: &str,
    property: &str,
    value: Value,
    ctx: &mut Context,
) -> Result<()> {
    super::properties::set_property(object_type, object_data, property, value, ctx)
}

pub fn dispatch_method_call(
    object_type: &str,
    object_data: &str,
    method: &str,
    args: &[Value],
) -> Result<Value> {
    super::methods::call_method(object_type, object_data, method, args)
}
