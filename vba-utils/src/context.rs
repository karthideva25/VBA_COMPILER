// vba-utils/src/context.rs

use std::collections::{HashMap, HashSet};
use crate::ast::Statement;
use crate::host::ComRegistry;
use crate::runtime_config::RuntimeConfig;

pub type VbaValue = Value;

/// VBA VarType constants - used by VarType() function
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum VarType {
    VbEmpty = 0,
    VbNull = 1,
    VbInteger = 2,
    VbLong = 3,
    VbSingle = 4,
    VbDouble = 5,
    VbCurrency = 6,
    VbDate = 7,
    VbString = 8,
    VbObject = 9,
    VbError = 10,
    VbBoolean = 11,
    VbVariant = 12,
    VbByte = 17,
    VbLongLong = 20,
    VbUserDefinedType = 36,
    VbArray = 8192,  // OR'ed with base type for arrays
}

impl VarType {
    /// Get the numeric value for VarType()
    pub fn as_i64(&self) -> i64 {
        match self {
            VarType::VbEmpty => 0,
            VarType::VbNull => 1,
            VarType::VbInteger => 2,
            VarType::VbLong => 3,
            VarType::VbSingle => 4,
            VarType::VbDouble => 5,
            VarType::VbCurrency => 6,
            VarType::VbDate => 7,
            VarType::VbString => 8,
            VarType::VbObject => 9,
            VarType::VbError => 10,
            VarType::VbBoolean => 11,
            VarType::VbVariant => 12,
            VarType::VbByte => 17,
            VarType::VbLongLong => 20,
            VarType::VbUserDefinedType => 36,
            VarType::VbArray => 8192,
        }
    }
    
    /// Get the type name for TypeName() function
    pub fn type_name(&self) -> &'static str {
        match self {
            VarType::VbEmpty => "Empty",
            VarType::VbNull => "Null",
            VarType::VbInteger => "Integer",
            VarType::VbLong => "Long",
            VarType::VbSingle => "Single",
            VarType::VbDouble => "Double",
            VarType::VbCurrency => "Currency",
            VarType::VbDate => "Date",
            VarType::VbString => "String",
            VarType::VbObject => "Object",
            VarType::VbError => "Error",
            VarType::VbBoolean => "Boolean",
            VarType::VbVariant => "Variant",
            VarType::VbByte => "Byte",
            VarType::VbLongLong => "LongLong",
            VarType::VbUserDefinedType => "UserDefinedType",
            VarType::VbArray => "Array",
        }
    }
}

/// A runtime VBA value: either integer or string.
#[derive(Debug, Clone)]
pub enum Value {
    Boolean(bool),
    Byte(u8),
    Currency(f64),
    Date(chrono::NaiveDate),
    DateTime(chrono::NaiveDateTime),  // Date + Time combined (for Now())
    Time(chrono::NaiveTime),           // Time only (for Time())
    Double(f64),
    Decimal(f64), 
    Integer(i64),
    Long(i32),         // new: 32-bit signed
    LongLong(i64),     // new: 64-bit signed
    Object(Option<Box<Value>>), 
    Single(f32), 
    String(String),
    Empty,                      // Uninitialized Variant
    Null,                       // Database NULL (optional)
    Error(i32),                 // VBA Error value (CVErr result)
    UserType { 
        type_name: String,
        fields: HashMap<String, Value>,
    },
}

impl Value {
    pub fn as_string(&self) -> String {
        match self {
            Value::Integer(i) => i.to_string(),
            Value::Long(l) => l.to_string(),
            Value::LongLong(ll) => ll.to_string(),
            Value::Byte(b)    => b.to_string(),
            Value::String(s)  => s.clone(),
            Value::Boolean(b) => b.to_string(),
            Value::Currency(c) => format!("{:.4}", c),
            Value::Date(d) => d.format("%m/%d/%Y").to_string(),
            Value::DateTime(dt) => dt.format("%m/%d/%Y %H:%M:%S").to_string(),
            Value::Time(t) => t.format("%H:%M:%S").to_string(),
            Value::Double(f)  => f.to_string(),
            Value::Decimal(f) => f.to_string(),
            Value::Object(None) => "Nothing".into(),
            Value::Object(Some(inner)) => inner.as_string(),   
            Value::Single(s) => s.to_string(), 
            Value::UserType { type_name, .. } => { 
                format!("<{} instance>", type_name)
            }
            Value::Empty => String::new(),
            Value::Null => "Null".into(),
            Value::Error(e) => format!("Error {}", e),
        }
    }
    
    pub fn as_integer(&self) -> Option<i64> {
        match self {
            Value::Boolean(b) => Some(if *b { 1 } else { 0 }),
            Value::Byte(b)    => Some(*b as i64),  // Convert byte to i64
            Value::Currency(c) => Some(*c as i64),
            Value::Date(_) => None,
            Value::DateTime(_) => None,
            Value::Time(_) => None,
            Value::Double(f)  => Some(*f as i64),
            Value::Decimal(f) => Some(*f as i64),
            Value::Integer(i) => Some(*i),
            Value::Long(l) => Some(*l as i64),
            Value::LongLong(ll) => Some(*ll),
            Value::Object(None) => None, // ✅ new: Nothing -> None
            Value::Object(Some(inner)) => inner.as_integer(), // ✅ delegate to inner
            Value::Single(f) => Some(*f as i64), // ✅ new: Single
            Value::String(s)  => s.parse::<i64>().ok(),
            Value::UserType { .. } => None,
            Value::Empty => Some(0),  // ✅ Empty converts to 0 in numeric context
            Value::Null => None,
            Value::Error(e) => Some(*e as i64),
        }
    }
    // Get a field value from a user-defined type
    pub fn get_field(&self, field_name: &str) -> Option<Value> {
        match self {
            Value::UserType { fields, .. } => fields.get(field_name).cloned(),
            _ => None,
        }
    }
    
    /// Set a field value in a user-defined type
    pub fn set_field(&mut self, field_name: &str, value: Value) -> Result<(), String> {
        match self {
            Value::UserType { fields, .. } => {
                fields.insert(field_name.to_string(), value);
                Ok(())
            }
            _ => Err(format!("Cannot set field '{}' on non-UserType value", field_name)),
        }
    }
    
    /// Check if this value is a user-defined type
    pub fn is_user_type(&self) -> bool {
        matches!(self, Value::UserType { .. })
    }
    
    /// Get the type name if this is a user-defined type
    pub fn get_type_name(&self) -> Option<&str> {
        match self {
            Value::UserType { type_name, .. } => Some(type_name.as_str()),
            _ => None,
        }
    }
    
    /// Get all field names for a user-defined type
    pub fn get_field_names(&self) -> Option<Vec<String>> {
        match self {
            Value::UserType { fields, .. } => {
                Some(fields.keys().cloned().collect())
            }
            _ => None,
        }
    }

    /// Get the VarType of this value (for VarType() function)
    pub fn get_var_type(&self) -> VarType {
        match self {
            Value::Empty => VarType::VbEmpty,
            Value::Null => VarType::VbNull,
            Value::Boolean(_) => VarType::VbBoolean,
            Value::Byte(_) => VarType::VbByte,
            Value::Integer(_) => VarType::VbInteger,
            Value::Long(_) => VarType::VbLong,
            Value::LongLong(_) => VarType::VbLongLong,
            Value::Single(_) => VarType::VbSingle,
            Value::Double(_) => VarType::VbDouble,
            Value::Decimal(_) => VarType::VbDouble, // Decimal treated as Double
            Value::Currency(_) => VarType::VbCurrency,
            Value::Date(_) => VarType::VbDate,
            Value::DateTime(_) => VarType::VbDate,  // DateTime is also vbDate
            Value::Time(_) => VarType::VbDate,      // Time is also vbDate in VBA
            Value::String(_) => VarType::VbString,
            Value::Object(_) => VarType::VbObject,
            Value::Error(_) => VarType::VbError,
            Value::UserType { .. } => VarType::VbUserDefinedType,
        }
    }
    
    /// Check if this value is Empty
    pub fn is_empty(&self) -> bool {
        matches!(self, Value::Empty)
    }
    
    /// Check if this value is Null
    pub fn is_null(&self) -> bool {
        matches!(self, Value::Null)
    }
    
    /// Get the type name string (for TypeName() function)
    pub fn type_name(&self) -> String {
        match self {
            Value::UserType { type_name, .. } => type_name.clone(),
            _ => self.get_var_type().type_name().to_string(),
        }
    }

}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum DeclaredType {
    Boolean,
    Byte,
    Currency,
    Date,
    Double,
    Decimal, //not currently supported in excel
    Integer,
    Long,       // new
    LongLong,   // new
    Object,
    Single,
    String,
    Variant, // when no type is provided in Dim
}

impl DeclaredType {
    pub fn from_opt_str(s: Option<&str>) -> Self {
        match s.map(|t| t.trim().to_ascii_lowercase()).as_deref() {
            Some("byte")     => DeclaredType::Byte,
            Some("integer")  => DeclaredType::Integer,
            Some("currency") => DeclaredType::Currency,
            Some("date")     => DeclaredType::Date,
            Some("double")   => DeclaredType::Double,
            Some("decimal")  => DeclaredType::Decimal,
            Some("string")   => DeclaredType::String,
            Some("boolean")  => DeclaredType::Boolean,
            _                => DeclaredType::Variant,
        }
    }
}

/// What kind of scope we’re pushing (purely for debugging/trace output).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ScopeKind {
    Subroutine,
    Function,
    Block, // If/For/While/etc.
}

impl Default for ScopeKind {
    fn default() -> Self {
        ScopeKind::Block
    }
}

#[derive(Debug, Default, Clone)]
struct ScopeFrame {
    name: Option<String>,
    kind: ScopeKind,
    vars: HashMap<String, Value>,
    types: HashMap<String, DeclaredType>,
}

/// Execution context: holds variables, output **and** subroutine definitions.
///
/// NOTE: `variables` remains your **global** scope for backward compatibility.
/// New local scopes are held in the private `scopes` stack.
#[derive(Debug, Default)]
pub struct Context {
    /// Messages logged (e.g. via MsgBox)
    pub output: Vec<String>,
    /// Global/module-level variables (backward compatible)
    pub variables: HashMap<String, Value>,
    /// Subroutine/Function definitions: name → (params, body)
    pub subs: HashMap<String, (Vec<crate::ast::Parameter>, Vec<Statement>)>,
    /// Function return types: name → return type
    pub function_return_types: HashMap<String, Option<String>>,

    // global declared types (module level), parallel to `variables`
    global_types: HashMap<String, DeclaredType>,
    
    pub enums: HashMap<String, EnumDefinition>,

    pub types: HashMap<String, TypeDefinition>,

    // private overlay scopes (top is current). Not visible to callers.
    scopes: Vec<ScopeFrame>,

    pub err: Option<ErrObject>,          // last runtime error
    pub on_error_mode: OnErrorMode,      // current mode
    pub on_error_label: Option<String>,  // target label if mode == GoTo
    pub resume_valid: bool,
    pub resume_pc: Option<usize>,
    pub resume_location: Option<ResumeLocation>,

    pub option_explicit: bool,           // Whether Option Explicit is active
    declared_vars: HashSet<String>,
    pub com_registry: ComRegistry,
    
    /// Stack of With block objects (for .Property syntax)
    pub with_stack: Vec<Value>,
    
    /// Runtime configuration (timezone, locale, workbook, user)
    /// Passed from application layer at session start
    pub runtime_config: RuntimeConfig,
}

impl Context {
    pub fn has_sub(&self, name: &str) -> bool {
        self.subs.contains_key(name)
    }
    /// Register a subroutine (called during Phase 1)
    pub fn register_sub(&mut self, name: &str, params: &[crate::ast::Parameter], body: &[Statement]) {
        self.subs.insert(
            name.to_string(),
            (params.to_vec(), body.to_vec()),
        );
    }

    /// Register a function (called during Phase 1)
    pub fn register_function(
        &mut self,
        name: &str,
        params: &[crate::ast::Parameter],
        body: &[Statement],
        return_type: &Option<String>,
    ) {
        self.subs.insert(
            name.to_string(),
            (params.to_vec(), body.to_vec()),
        );
        self.function_return_types.insert(name.to_string(), return_type.clone());
    }

    /// Register a property (called during Phase 1)
    pub fn register_property(
        &mut self,
        prop_type: &str, // "Get", "Let", "Set"
        name: &str,
        params: &[crate::ast::Parameter],
        body: &[Statement],
    ) {
        // Store with type prefix to distinguish Get/Let/Set
        let key = format!("{}_{}", prop_type, name);
        self.register_sub(&key, params, body);
    }

    pub fn log(&mut self, msg: &str) {
        println!("{}", msg);
        self.output.push(msg.to_string());
    }

    /// Back-compat assignment:
    /// - If a variable already exists in any active scope (from innermost to outermost), update it there.
    /// - Otherwise, assign to the **global** map (as the old code did).
    pub fn set_var(&mut self, name: String, val: Value) {
        // Try innermost → outermost local scopes
        for i in (0..self.scopes.len()).rev() {
            if self.scopes[i].vars.contains_key(&name) {
                self.scopes[i].vars.insert(name, val);
                return;
            }
        }
        // Fall back to global (old behavior)
        self.variables.insert(name, val);
    }
    pub fn set_var_type(&mut self, name: String, ty: DeclaredType) {
        // try innermost → outermost local scopes
        for i in (0..self.scopes.len()).rev() {
            if self.scopes[i].vars.contains_key(&name) || self.scopes[i].types.contains_key(&name) {
                self.scopes[i].types.insert(name, ty);
                return;
            }
        }
        // otherwise mark as global type
        self.global_types.insert(name, ty);
    }


    /// Scope-aware lookup: check innermost → outermost local scopes first, then global.
    pub fn get_var(&self, name: &str) -> Option<Value> {
        for frame in self.scopes.iter().rev() {
            if let Some(v) = frame.vars.get(name) {
                return Some(v.clone());
            }
        }
        self.variables.get(name).cloned()
    }
    pub fn get_var_type(&self, name: &str) -> Option<DeclaredType> {
        for frame in self.scopes.iter().rev() {
            if let Some(t) = frame.types.get(name) {
                return Some(*t);
            }
        }
        self.global_types.get(name).copied()
    }


    /// Define a subroutine for later calls.
    pub fn define_sub(&mut self, name: String, params: Vec<crate::ast::Parameter>, body: Vec<Statement>) {
        self.subs.insert(name, (params, body));
    }

    /// Define a function with return type.
    pub fn define_function(&mut self, name: String, params: Vec<crate::ast::Parameter>, body: Vec<Statement>, return_type: Option<String>) {
        self.subs.insert(name.clone(), (params, body));
        self.function_return_types.insert(name, return_type);
    }

    /// Check if a name is a function (has return type registered)
    pub fn is_function(&self, name: &str) -> bool {
        self.function_return_types.contains_key(name)
    }

    /// Save/restore **global** variable scope (unchanged API & semantics).
    /// If you used this around sub calls before, it will continue to work.
    pub fn save_scope(&self) -> HashMap<String, Value> {
        self.variables.clone()
    }
    pub fn restore_scope(&mut self, old: HashMap<String, Value>) {
        self.variables = old;
    }

    // === NEW: Scope management (non-breaking additions) =====================

    /// Push a new local scope on the stack.
    pub fn push_scope(&mut self, name: impl Into<String>, kind: ScopeKind) {
        self.scopes.push(ScopeFrame {
            name: Some(name.into()),
            kind,
            vars: HashMap::new(),
            types: HashMap::new(),
        });
    }

    /// Pop the current local scope. No-op if there is none.
    pub fn pop_scope(&mut self) {
        let _ = self.scopes.pop();
    }

    /// Declare a local (or parameter) in the current scope. If no scope is active,
    /// declares in global (so callers don’t have to special-case).
    pub fn declare_local(&mut self, name: impl Into<String>, initial: Value) {
        if let Some(top) = self.scopes.last_mut() {
            top.vars.insert(name.into(), initial);
        } else {
            // No active local scope, fall back to global for back-compat.
            self.variables.insert(name.into(), initial);
        }
    }

    /// Helper: run a block within a scope (ensures pop even on early return/err).
    pub fn with_scope<R, F>(&mut self, name: impl Into<String>, kind: ScopeKind, f: F) -> R
    where
        F: FnOnce(&mut Context) -> R,
    {
        self.push_scope(name, kind);
        // In real interpreter code you may want error handling here.
        let r = f(self);
        self.pop_scope();
        r
    }

    /// (Optional) Full snapshot/restore if you want to freeze locals+globals.
    /// Keeping separate from old save/restore avoids breaking the API.
    pub fn save_all_scopes(&self) -> SavedScopes {
        SavedScopes {
            globals: self.variables.clone(),
            stack: self.scopes.iter().map(|f| SavedScopeFrame {
                name: f.name.clone(),
                kind: f.kind,
                vars: f.vars.clone(),
                types: f.types.clone(),
            }).collect(),
        }
    }

    pub fn restore_all_scopes(&mut self, snap: SavedScopes) {
        self.variables = snap.globals;
        self.scopes = snap.stack.into_iter().map(|f| ScopeFrame {
            name: f.name,
            kind: f.kind,
            vars: f.vars,
            types: f.types,
        }).collect();
    }

    // Add method to define an enum:
    pub fn define_enum(&mut self, name: String, members: HashMap<String, i64>) {
        self.enums.insert(name.clone(), EnumDefinition {
            name,
            members,
        });
    }
    
    // Add method to get enum member value:
    pub fn get_enum_value(&self, enum_name: &str, member_name: &str) -> Option<i64> {
        self.enums.get(enum_name)
            .and_then(|enum_def| enum_def.members.get(member_name))
            .copied()
    }
    
    // Add method to resolve qualified enum reference (e.g., SecurityLevel.SecurityLevel1)
    pub fn resolve_enum_member(&self, qualified_name: &str) -> Option<Value> {
        // Split on dot to get enum_name.member_name
        if let Some(dot_pos) = qualified_name.find('.') {
            let enum_name = &qualified_name[..dot_pos];
            let member_name = &qualified_name[dot_pos + 1..];
            
            if let Some(value) = self.get_enum_value(enum_name, member_name) {
                return Some(Value::Integer(value));
            }
        }
        None
    }

    // Add method to define a type:
    pub fn define_type(&mut self, name: String, fields: HashMap<String, FieldDefinition>) {
        self.types.insert(name.clone(), TypeDefinition {
            name,
            fields,
        });
    }
    
    // Add method to get type definition:
    pub fn get_type_definition(&self, type_name: &str) -> Option<&TypeDefinition> {
        self.types.get(type_name)
    }
    
    // Add method to check if a type is defined:
    pub fn is_type_defined(&self, type_name: &str) -> bool {
        self.types.contains_key(type_name)
    }
    
    // Add method to create an instance of a user-defined type
    pub fn create_type_instance(&self, type_name: &str) -> Option<Value> {
        let type_def = self.get_type_definition(type_name)?;
        let mut fields = HashMap::new();
        
        // Initialize all fields with default values
        for (field_name, field_def) in &type_def.fields {
            let default_value = match field_def.field_type.as_str() {
                "Integer" | "Long" | "Byte" => Value::Integer(0),
                "String" => Value::String(String::new()),
                "Boolean" => Value::Boolean(false),
                _ => Value::String(String::new()),  // Default for unknown types
            };
            fields.insert(field_name.clone(), default_value);
        }
        
        Some(Value::UserType {
            type_name: type_name.to_string(),
            fields,
        })
    }
    pub fn list_all_vars(&self) -> Vec<String> {
        let mut vars = Vec::new();
        
        // Get local variables from current scope (if any)
        if let Some(current_scope) = self.scopes.last() {
            for name in current_scope.vars.keys() {
                vars.push(format!("Local: {}", name));
            }
        }
        
        // Get global variables
        for name in self.variables.keys() {
            vars.push(format!("Global: {}", name));
        }
        
        vars
    }
    /// Alternative: more detailed version with values
    pub fn debug_vars(&self) -> String {
        let mut output = String::new();
        
        output.push_str("=== Variables ===\n");
        
        // Local variables from current scope
        if let Some(current_scope) = self.scopes.last() {
            if !current_scope.vars.is_empty() {
                output.push_str("Local scope:\n");
                for (name, value) in &current_scope.vars {
                    output.push_str(&format!("  {} = {:?}\n", name, value));
                }
            }
        }
        
        // Global variables
        if !self.variables.is_empty() {
            output.push_str("Global scope:\n");
            for (name, value) in &self.variables {
                output.push_str(&format!("  {} = {:?}\n", name, value));
            }
        }
        
        if self.scopes.is_empty() && self.variables.is_empty() {
            output.push_str("  (no variables)\n");
        }
        
        output
    }

    pub fn enable_option_explicit(&mut self) {
        self.option_explicit = true;
        self.log("Option Explicit enabled - all variables must be declared");
    }
    
    /// Check if Option Explicit is enabled
    pub fn is_option_explicit(&self) -> bool {
        self.option_explicit
    }
    
    /// Mark a variable as declared (for Option Explicit checking)
    pub fn declare_variable(&mut self, name: &str) {
        self.declared_vars.insert(name.to_string());
    }
    
    /// Check if a variable has been declared
    pub fn is_variable_declared(&self, name: &str) -> bool {
        self.declared_vars.contains(name)
    }
    
    /// Validate variable usage when Option Explicit is enabled
    pub fn validate_variable_usage(&self, name: &str) -> Result<(), String> {
        if self.option_explicit && !self.is_variable_declared(name) {
            Err(format!("Variable '{}' is used without being declared. Use Dim to declare it first.", name))
        } else {
            Ok(())
        }
    }
    
    pub fn new() -> Self {
        Self::with_config(RuntimeConfig::default())
    }
    
    /// Create a new Context with the given runtime configuration.
    /// 
    /// This is the preferred way to create a Context in production:
    /// ```rust,ignore
    /// let config = RuntimeConfig::builder()
    ///     .timezone("Asia/Kolkata")
    ///     .workbook_id("wb-12345")
    ///     .user_id("user-67890")
    ///     .build();
    /// let mut ctx = Context::with_config(config);
    /// ```
    pub fn with_config(config: RuntimeConfig) -> Self {
        Self {
            variables: HashMap::new(),
            scopes: Vec::new(),
            subs: HashMap::new(),
            function_return_types: HashMap::new(),
            types: HashMap::new(),
            enums: HashMap::new(),
            global_types: HashMap::new(),
            declared_vars: HashSet::new(),
            option_explicit: false,
            on_error_mode: OnErrorMode::None,
            on_error_label: None, 
            err: None,
            resume_valid: false,
            resume_pc: None,
            resume_location: None,
            output: Vec::new(),
            com_registry: ComRegistry::new(),
            with_stack: Vec::new(),
            runtime_config: config,
        }
    }

}

/// Full-snapshot types are private by default; make them `pub` if you need them externally.
#[derive(Debug, Clone)]
pub struct SavedScopes {
    globals: HashMap<String, Value>,
    stack: Vec<SavedScopeFrame>,
}

#[derive(Debug, Clone)]
struct SavedScopeFrame {
    name: Option<String>,
    kind: ScopeKind,
    vars: HashMap<String, Value>,
    types: HashMap<String, DeclaredType>,
}
// === Error handling state (VBA-style) =====================================

#[derive(Debug, Clone, Default)]
pub struct ErrObject {
    pub number: i32,
    pub description: String,
    pub source: String, 
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OnErrorMode {
    None,       // default: no handler → unhandled error stops the Sub
    ResumeNextAuto, // skip failing statement, continue at next
    GoTo,       // jump to label
}
#[derive(Debug, Clone, Default)]
pub struct ResumeLocation {
    pub frame_id: usize,    // CHANGED: now frame ID, not list_id
    pub pc: usize,
    pub parent_pc: Option<usize>,
}
impl Default for OnErrorMode {
    fn default() -> Self { OnErrorMode::None }
}

#[derive(Debug)]
pub struct ProcHandlerState {
    pub on_error_mode: OnErrorMode,
    pub on_error_label: Option<String>, // valid when mode==GoTo

    pub err: Option<ErrObject>,

    // Saved program counters for Resume/Resume Next (per-fault)
    pub resume_pc: Option<usize>,     // index of the faulting statement
    pub resume_valid: bool,           // set when inside a handler block reached by GoTo
}

// Add enum definition structure:
#[derive(Debug, Clone)]
pub struct EnumDefinition {
    pub name: String,
    pub members: HashMap<String, i64>,  // member_name -> value
}
// Add type definition structure:
#[derive(Debug, Clone)]
pub struct TypeDefinition {
    pub name: String,
    pub fields: HashMap<String, FieldDefinition>,
}

#[derive(Debug, Clone)]
pub struct FieldDefinition {
    pub name: String,
    pub field_type: String,
    pub string_length: Option<i64>,
    pub is_array: bool,
}