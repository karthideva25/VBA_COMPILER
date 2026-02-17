// src/host/excel/properties/application/mod.rs
// Property handlers for Application object

pub mod interaction;
pub mod calculation;
pub mod metadata;
pub mod events;
pub mod references;

use anyhow::Result;
use crate::context::{Context, Value};

/// Route property get requests to specialized handlers
pub fn get_property(property: &str, _ctx: &mut Context) -> Result<Value> {
    match property.to_lowercase().as_str() {
        // Interaction properties
        "displayalerts" => interaction::get_property(property),
        "screenupdating" => interaction::get_property(property),
        "enableevents" => interaction::get_property(property),
        
        // Calculation properties
        "calculation" => calculation::get_property(property),
        
        // Metadata properties
        "username" | "useremailid" | "creatorname" | "creatoremailid" => metadata::get_property(property),
        
        // Event handlers
        "oncalculate" | "ondata" | "ondoubleclick" | "onentry" | "onsheetactivate" | "onsheetdeactivate" => events::get_property(property),
        
        // Reference properties
        "referencestyle" | "cutcopymode" => references::get_property(property),
        
        _ => Err(anyhow::anyhow!("Unknown Application property: {}", property)),
    }
}

/// Route property set requests to specialized handlers
pub fn set_property(property: &str, value: Value, _ctx: &mut Context) -> Result<()> {
    match property.to_lowercase().as_str() {
        "displayalerts" => interaction::set_property(property, value),
        "screenupdating" => interaction::set_property(property, value),
        "enableevents" => interaction::set_property(property, value),
        "calculation" => calculation::set_property(property, value),
        "username" | "useremailid" | "creatorname" | "creatoremailid" => metadata::set_property(property, value),
        "oncalculate" | "ondata" | "ondoubleclick" | "onentry" | "onsheetactivate" | "onsheetdeactivate" => events::set_property(property, value),
        "referencestyle" | "cutcopymode" => references::set_property(property, value),
        _ => Err(anyhow::anyhow!("Cannot set Application property: {}", property)),
    }
}
