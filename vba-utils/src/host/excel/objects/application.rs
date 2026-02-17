use std::collections::HashMap;
use anyhow::Result;
use crate::context::{Context, Value};
use crate::host::ComObject;

/// Excel Application object - root object for Excel automation
#[derive(Debug)]
pub struct ExcelApplication {
    // Display & Interaction Properties
    pub display_alerts: bool,
    pub screen_updating: bool,
    pub enable_events: bool,
    
    // Calculation Properties
    pub calculation: String,
    
    // Reference Style
    pub reference_style: i32, // 1 = A1, 2 = R1C1
    
    // Cut/Copy Mode
    pub cut_copy_mode: i32,
    
    // User Information
    pub user_name: String,
    pub user_email_id: String,
    pub creator_name: String,
    pub creator_email_id: String,
    
    // Event Handlers
    pub on_calculate: String,
    pub on_data: String,
    pub on_double_click: String,
    pub on_entry: String,
    pub on_sheet_activate: String,
    pub on_sheet_deactivate: String,
    
    // Custom Properties
    pub custom_properties: HashMap<String, String>,
}

impl ExcelApplication {
    pub fn new() -> Self {
        Self {
            display_alerts: true,
            screen_updating: true,
            enable_events: true,
            calculation: "Automatic".to_string(),
            reference_style: 1, // A1 style
            cut_copy_mode: 0,
            user_name: "User".to_string(),
            user_email_id: String::new(),
            creator_name: String::new(),
            creator_email_id: String::new(),
            on_calculate: String::new(),
            on_data: String::new(),
            on_double_click: String::new(),
            on_entry: String::new(),
            on_sheet_activate: String::new(),
            on_sheet_deactivate: String::new(),
            custom_properties: HashMap::new(),
        }
    }
}

impl Default for ExcelApplication {
    fn default() -> Self {
        Self::new()
    }
}

/// Implement ComObject trait for Application
impl ComObject for ExcelApplication {
    fn get_property(&self, name: &str, ctx: &mut Context) -> Result<Value> {
        super::super::properties::application::get_property(name, ctx)
    }

    fn set_property(&mut self, name: &str, value: Value, ctx: &mut Context) -> Result<()> {
        super::super::properties::application::set_property(name, value, ctx)
    }

    fn call_method(&mut self, name: &str, args: &[Value], ctx: &mut Context) -> Result<Value> {
        super::super::methods::application::call_method(name, args, ctx)
    }

    fn type_name(&self) -> &str {
        "Application"
    }
}
