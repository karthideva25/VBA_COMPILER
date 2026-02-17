// src/host/mod.rs

pub mod excel;

use std::cell::RefCell;
use std::collections::HashMap;
use std::rc::Rc;

use anyhow::{anyhow, Result};

use crate::context::{Context, Value};

/// Trait implemented by all COM-style host objects.
pub trait ComObject {
    fn get_property(&self, name: &str, ctx: &mut Context) -> Result<Value>;
    fn set_property(&mut self, name: &str, value: Value, ctx: &mut Context) -> Result<()>;
    fn call_method(&mut self, name: &str, args: &[Value], ctx: &mut Context) -> Result<Value>;
    fn type_name(&self) -> &str;
}

pub type ComObjectHandle = Rc<RefCell<dyn ComObject>>;

/// Registry of COM objects (Application, Range, Workbook, etc.)
pub struct ComRegistry {
    globals: HashMap<String, ComObjectHandle>,
}

impl ComRegistry {
    pub fn new() -> Self {
        Self {
            globals: HashMap::new(),
        }
    }

    /// Register an instance and return its ID.
    pub fn register_instance(&mut self, _obj: ComObjectHandle) -> usize {
        // For now, use a simple static counter
        static COUNTER: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(1);
        COUNTER.fetch_add(1, std::sync::atomic::Ordering::SeqCst)
    }

    /// Register a named global COM object (e.g. "Application").
    pub fn register_global(&mut self, name: impl Into<String>, obj: ComObjectHandle) {
        self.globals.insert(name.into(), obj);
    }

    /// Look up a previously registered global object.
    pub fn get_global(&self, name: &str) -> Option<ComObjectHandle> {
        self.globals.get(name).cloned()
    }
}

impl Default for ComRegistry {
    fn default() -> Self {
        Self::new()
    }
}

impl std::fmt::Debug for ComRegistry {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.debug_struct("ComRegistry")
            .field("globals_len", &self.globals.len())
            .finish()
    }
}

/// Common dispatch helper used by the interpreter for COM property/method calls.
///
/// - `object_name`: name of the registered COM object (e.g. "Application")
/// - `property_or_method`: property or method name (e.g. "DisplayAlerts", "Quit")
/// - `args`: optional arguments
/// - `is_set`: true for property set, false for get/call
pub fn dispatch_com_call(
    object_name: &str,
    property_or_method: &str,
    args: Option<&[Value]>,
    is_set: bool,
    ctx: &mut Context,
) -> Result<Value> {
    let handle = ctx
        .com_registry
        .get_global(object_name)
        .ok_or_else(|| anyhow!("COM object '{}' not registered", object_name))?;

    let mut borrowed = handle
        .try_borrow_mut()
        .map_err(|_| anyhow!("COM object '{}' is already borrowed", object_name))?;

    if is_set {
        // Property set: exactly one argument required
        let args = args.ok_or_else(|| anyhow!("property set requires one argument"))?;
        if args.len() != 1 {
            return Err(anyhow!(
                "property set '{}.{}' expects 1 argument, got {}",
                object_name,
                property_or_method,
                args.len()
            ));
        }
        borrowed.set_property(property_or_method, args[0].clone(), ctx)?;
        Ok(Value::Empty)
    } else if let Some(args) = args {
        // Method call: arguments present
        borrowed.call_method(property_or_method, args, ctx)
    } else {
        // Property get: no arguments
        borrowed.get_property(property_or_method, ctx)
    }
}
