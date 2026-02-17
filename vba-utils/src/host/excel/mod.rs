// src/host/excel/mod.rs

pub mod engine;
pub mod static_engine;
pub mod properties;
pub mod methods;
pub mod objects;

use std::cell::RefCell;
use std::rc::Rc;

use crate::context::Context;
use crate::host::ComObjectHandle;

use self::objects::application::ExcelApplication;

/// Initialize the Excel host environment and register default COM objects.
pub fn initialize_excel_host(ctx: &mut Context) {
    // Initialize the Excel engine
    // Paths to resource files and app cache
    let resource_path = "/Users/poornema-13898/Downloads/SamplePOCMacro/resources";
    let local_path = "/Users/poornema-13898/Downloads/SamplePOCMacro/AppLocal";
    
    match engine::initialize_engine(resource_path, local_path) {
        Ok(_) => eprintln!("✅ Excel engine initialized"),
        Err(e) => eprintln!("⚠️  Failed to initialize Excel engine: {}", e),
    }
    
    // Register global Excel.Application
    let app: ComObjectHandle = Rc::new(RefCell::new(ExcelApplication::new()));
    ctx.com_registry.register_global("Application", app);

    // If you later want aliases like "Excel.Application", you can register them here
    // using ctx.com_registry.get_global("Application") and re-inserting.
}