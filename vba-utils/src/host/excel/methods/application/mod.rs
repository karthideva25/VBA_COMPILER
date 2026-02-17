// src/host/excel/methods/application/mod.rs
// Method handlers for Application object

pub mod utility;
pub mod calculation;
pub mod navigation;
pub mod interaction;

use anyhow::Result;
use crate::context::{Context, Value};

/// Route method calls to specialized handlers
pub fn call_method(method: &str, args: &[Value], _ctx: &mut Context) -> Result<Value> {
    match method.to_lowercase().as_str() {
        // Calculation methods
        "calculate" | "calculatefull" => calculation::call_method(method, args),
        
        // Navigation methods
        "goto" | "activateprevious" | "activatenext" => navigation::call_method(method, args),
        
        // Interaction methods
        "quit" | "wait" | "inputbox" | "ontime" => interaction::call_method(method, args),
        
        // Utility methods
        "centimeterstopo" | "inchestopoint" | "convertformula" | "evaluate" => utility::call_method(method, args),
        
        _ => Err(anyhow::anyhow!("Unknown Application method: {}", method)),
    }
}
