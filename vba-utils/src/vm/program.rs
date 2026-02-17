use crate::ast::{Program, Statement};
use crate::context::Context;
use crate::interpreter::{execute_statement, run_subroutine};

/// The main entry point for executing a VBA program.
/// Follows VBA's 3-phase execution model:
/// 1. Register all declarations (Types, Enums, Subs)
/// 2. Initialize module-level variables
/// 3. Execute an entrypoint (AutoOpen, Workbook_Open, etc.)
pub struct ProgramExecutor {
    program: Program,
}

impl ProgramExecutor {
    pub fn new(program: Program) -> Self {
        Self { program }
    }

    /// Execute the full 3-phase process with automatic entrypoint detection
    pub fn execute(&self, ctx: &mut Context) -> Result<(), String> {
        // Phase 1: Register declarations
        self.register_declarations(ctx)?;
        // Initialize Excel host
        crate::host::excel::initialize_excel_host(ctx);
        
        // Phase 2: Initialize module variables
        self.initialize_module_variables(ctx)?;

        // Phase 3: Run entrypoint (auto-detect)
        let entrypoint = self.detect_entrypoint(ctx);
        if let Some(name) = entrypoint {
            eprintln!("▶️ Auto-detected entrypoint: {}", name);
            // run_subroutine does not return Result, so no `?` here
            run_subroutine(ctx, &name);
        } else {
            eprintln!("⚠️ No entrypoint found (AutoOpen, Workbook_Open, Main)");
        }

        Ok(())
    }

    /// Execute with a specific entrypoint
    pub fn execute_entrypoint(&self, ctx: &mut Context, entrypoint: &str) -> Result<(), String> {
        // Phase 1: Register declarations
        self.register_declarations(ctx)?;

        // Phase 2: Initialize module variables
        self.initialize_module_variables(ctx)?;

        // Phase 3: Run specified entrypoint
        eprintln!("▶️ Running entrypoint: {}", entrypoint);
        run_subroutine(ctx, entrypoint);

        Ok(())
    }

    /// Phase 1: Register all module-level declarations
    /// Order: Option Explicit → Types → Enums → Variables (declare) → Subs
    fn register_declarations(&self, ctx: &mut Context) -> Result<(), String> {
        // eprintln!("📦 Phase 1: Registering module declarations");

        // 1.1: Option Explicit (if present)
        for stmt in &self.program.statements {
            if let Statement::OptionExplicit = stmt {
                ctx.enable_option_explicit();
                // eprintln!("   ✅ Option Explicit enabled");
            }
        }

        // 1.2: Register Types FIRST (other things may depend on them)
        for stmt in &self.program.statements {
            if let Statement::Type { .. } = stmt {
                // let execute_statement handle define_type / etc.
                execute_statement(stmt, ctx, 0);
                // eprintln!("   ✅ Registered Type: {}", name);
            }
        }

        // 1.3: Register Enums SECOND
        for stmt in &self.program.statements {
            if let Statement::Enum { .. } = stmt {
                execute_statement(stmt, ctx, 0);
                // eprintln!("   ✅ Registered Enum: {}", name);
            }
        }

        // 1.4: (Const support can be added later when you add a `Const` variant)
        // for stmt in &self.program.statements {
        //     if let Statement::Const { name, .. } = stmt {
        //         execute_statement(stmt, ctx, 0);
        //         eprintln!("   ✅ Registered Const: {}", name);
        //     }
        // }

        // 1.5: Declare module-level variables FOURTH (don't initialize yet)
        for stmt in &self.program.statements {
            if let Statement::Dim { names } = stmt {
                for (var_name, _) in names {
                    ctx.declare_variable(var_name);
                    // eprintln!("   ✅ Declared module variable: {}", var_name);
                }
            }
        }

        // 1.6: Register Subs FIFTH (your AST uses `Subroutine`)
        for stmt in &self.program.statements {
            if let Statement::Subroutine { name, params, body } = stmt {
                ctx.register_sub(name, params, body);
                // eprintln!("   ✅ Registered Subroutine: {}", name);
            }
        }

        // 1.7: Register Functions SIXTH
        for stmt in &self.program.statements {
            if let Statement::Function { name, params, return_type, body } = stmt {
                ctx.register_function(name, params, body, return_type);
            }
        }

        // 1.8: Register Properties SEVENTH
        for stmt in &self.program.statements {
            match stmt {
                Statement::PropertyGet { name, params, body, return_type } => {
                    ctx.register_property("Get", name, params, body);
                    if let Some(ref rt) = return_type {
                        ctx.function_return_types.insert(format!("Get_{}", name), Some(rt.clone()));
                    }
                }
                Statement::PropertyLet { name, params, body } => {
                    ctx.register_property("Let", name, params, body);
                }
                Statement::PropertySet { name, params, body } => {
                    ctx.register_property("Set", name, params, body);
                }
                _ => {}
            }
        }

        Ok(())
    }

    /// Phase 2: Initialize module-level variables with their default values
    fn initialize_module_variables(&self, ctx: &mut Context) -> Result<(), String> {
        // eprintln!("🔧 Phase 2: Initializing module variables");

        for stmt in &self.program.statements {
            if let Statement::Dim { names } = stmt {
                // Execute the Dim statement to create instances
                execute_statement(stmt, ctx, 0);

                for (_var_name, type_name) in names {
                    let _type_str = type_name.as_deref().unwrap_or("Variant");
                    // eprintln!("   ✅ Initialized: {} As {}", var_name, type_str);
                }
            }
        }

        Ok(())
    }

    /// Detect common VBA entrypoints in priority order
    fn detect_entrypoint(&self, ctx: &Context) -> Option<String> {
        let candidates = [
            "AutoOpen",      // Word - opens with document
            "AutoExec",      // Word - starts with Word
            "Workbook_Open", // Excel - workbook opens
            "Auto_Open",     // Excel legacy
            "Main",          // Generic entry point
        ];

        for name in candidates {
            if ctx.has_sub(name) {
                return Some(name.to_string());
            }
        }

        None
    }

    /// Optional helpers if you want them:

    /// Check if a specific entrypoint exists
    pub fn has_entrypoint(&self, ctx: &Context, name: &str) -> bool {
        ctx.has_sub(name)
    }

    /// Get list of all interesting entrypoints present
    pub fn list_entrypoints(&self, ctx: &Context) -> Vec<String> {
        let candidates = [
            "AutoOpen", "AutoExec", "AutoClose", "AutoNew",
            "Workbook_Open", "Workbook_Close", "Workbook_BeforeSave",
            "Auto_Open", "Auto_Close",
            "Main",
        ];

        candidates
            .iter()
            .copied()
            .filter(|name| ctx.has_sub(name))
            .map(String::from)
            .collect()
    }
}

/// A handle for host environments to trigger VBA callbacks
pub struct VbaRuntime {
    ctx: Context,
}

impl VbaRuntime {
    /// Create a new runtime with initialized context
    pub fn new(program: Program) -> Result<Self, String> {
        // You don't have `Context::new()`, you have `Default`
        let mut ctx = Context::default();
        let executor = ProgramExecutor::new(program);

        // Run Phase 1 & 2 only (don't auto-execute entrypoint)
        executor.register_declarations(&mut ctx)?;
        executor.initialize_module_variables(&mut ctx)?;

        Ok(Self { ctx })
    }

    /// Execute a specific entrypoint/callback
    pub fn call_sub(&mut self, name: &str) -> Result<(), String> {
        // eprintln!("🔔 Host calling: {}", name);
        // run_subroutine returns (), so just call and then return Ok(())
        run_subroutine(&mut self.ctx, name);
        Ok(())
    }

    /// Execute a function and get return value (future work)
    pub fn call_function(
        &mut self,
        _name: &str,
        _args: Vec<crate::context::Value>,
    ) -> Result<crate::context::Value, String> {
        // TODO: Implement function calls with arguments and return values
        // This requires extending run_subroutine / a new run_function API.
        unimplemented!("Function calls with return values not yet implemented")
    }

    /// Get a variable value (for host to read VBA state)
    pub fn get_variable(&self, name: &str) -> Option<crate::context::Value> {
        self.ctx.get_var(name)
    }

    /// Set a variable value (for host to inject values)
    pub fn set_variable(&mut self, name: &str, value: crate::context::Value) {
        self.ctx.set_var(name.to_string(), value);
    }

    /// Check if a callback exists
    pub fn has_callback(&self, name: &str) -> bool {
        self.ctx.has_sub(name)
    }

    /// Get mutable access to context (for advanced host integration)
    pub fn context_mut(&mut self) -> &mut Context {
        &mut self.ctx
    }
}
