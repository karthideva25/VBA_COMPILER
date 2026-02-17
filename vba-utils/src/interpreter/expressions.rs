use anyhow::{bail, Result};
use crate::ast::Expression;
use crate::context::{Context, Value};
use super::builtins::{resolve_builtin_identifier};

pub(crate) fn evaluate_expression(expr: &Expression, ctx: &mut Context) -> Result<Value> {
    use Expression::*;

    match expr {
        // ——— Literals
        Integer(n) => Ok(Value::Integer(*n)),
        Byte(b) => Ok(Value::Byte(*b)),
        String(s)  => Ok(Value::String(s.clone())),
        Boolean(b) => Ok(Value::Boolean(*b)),
        Double(f)  => Ok(Value::Double(*f)),
        Decimal(f) => Ok(Value::Decimal(*f)),
        Single(s) => Ok(Value::Single(*s)),
        Currency(c) => Ok(Value::Currency(*c)),
        Date(d)     => Ok(Value::Date(*d)),

        // ——— Identifiers: built-in constants first, then variables
        Identifier(name) => {
            if name.eq_ignore_ascii_case("Err") {
                // Return a special marker or object representing Err
                // Since VBA's Err is an object, we can't return its properties directly
                // We need to return something that PropertyAccess can work with
                
                // Option 1: Return a special Value variant
                return Ok(Value::Object(Some(Box::new(Value::String("__ERR_OBJECT__".into())))));
                
                // OR Option 2: Just succeed and let PropertyAccess handle it
                // This is tricky because Identifier alone shouldn't fail
            }
            
            // 0. Handle special Excel global objects
            let name_lower = name.to_lowercase();
            if name_lower == "activesheet" {
                // Return a worksheet object reference
                return Ok(Value::Object(Some(Box::new(Value::String("ActiveSheet".into())))));
            }
            if name_lower == "activeworkbook" {
                return Ok(Value::Object(Some(Box::new(Value::String("ActiveWorkbook".into())))));
            }
            if name_lower == "application" {
                return Ok(Value::Object(Some(Box::new(Value::String("Application".into())))));
            }
            
            // 1. Check built-in constants first (vbTrue, vbCrLf, etc.)
            if let Some(v) = resolve_builtin_identifier(name) {
                return Ok(v);
            }
            
            // 2. Check if it's a qualified enum reference (EnumName.MemberName)
            //    This handles cases where the parser stored it as a single identifier
            if let Some(value) = ctx.resolve_enum_member(name) {
                return Ok(value);
            }
            
            // 3. Check Option Explicit before reading variable
            if let Err(e) = ctx.validate_variable_usage(name) {
                return Err(anyhow::anyhow!("{}", e));
            }
            
            // 4. Regular variable lookup
            ctx.get_var(name)
                .ok_or_else(|| anyhow::anyhow!("Variable '{}' not found", name))
        }
        
        BuiltInConstant(name) => {
            resolve_builtin_identifier(name)
                .ok_or_else(|| anyhow::anyhow!("Unknown built-in constant: {}", name))
        }

        // ——— Unary: op is a String (e.g., "+", "-", "Not")
        UnaryOp { op, expr } => {
            let v = evaluate_expression(expr, ctx)?;
            crate::interpreter::operations::eval_unary(op.as_str(), v)
        }

        // ——— Binary: op is a String (e.g., "+", "*", "<>", etc.)
        BinaryOp { left: lhs, op, right: rhs } => {
            // eprintln!("🔍 DEBUG: BinaryOp op={}, left={:?}, right={:?}", op, lhs, rhs);
            // Evaluate children first
            let l = evaluate_expression(lhs, ctx)?;
            let r = evaluate_expression(rhs, ctx)?;
        
            // Delegate to pure ops
            crate::interpreter::operations::eval_binary(ctx, op, l, r)
        }

        // ——— Function calls used as expressions
        FunctionCall { function, args } => {
            // Handle method calls on objects: obj.Method(args)
            // e.g., ws.Range("A1") where ws is a variable holding an object
            if let Expression::PropertyAccess { obj, property: method_name } = &**function {
                // Handle Err.Raise, Err.Clear as method calls with arguments
                if let Expression::Identifier(var_name) = &**obj {
                    if var_name.eq_ignore_ascii_case("Err") {
                        // Dispatch to Err method handler
                        if let Some(result) = crate::interpreter::builtins::handle_err_method(method_name, args, ctx)? {
                            return Ok(result);
                        }
                    }
                }
                
                // Evaluate the object to see what it is
                if let Expression::Identifier(var_name) = &**obj {
                    // Check if this variable holds an object reference
                    if let Some(var_val) = ctx.get_var(var_name) {
                        if let Value::Object(Some(inner)) = var_val {
                            if let Value::String(obj_type) = *inner {
                                // Object variable - dispatch method call
                                if obj_type == "ActiveSheet" && method_name.eq_ignore_ascii_case("Range") {
                                    // ws.Range("A1") where ws = ActiveSheet
                                    if let Some(first_arg) = args.first() {
                                        let address = evaluate_expression(first_arg, ctx)?;
                                        if let Value::String(addr) = address {
                                            return Ok(Value::Object(Some(Box::new(Value::String(format!("Range:{}", addr))))));
                                        }
                                    }
                                }
                            }
                        }
                    }
                    // Also handle direct ActiveSheet.Range(), Worksheets("...").Range(), etc.
                    if var_name.eq_ignore_ascii_case("ActiveSheet") && method_name.eq_ignore_ascii_case("Range") {
                        if let Some(first_arg) = args.first() {
                            let address = evaluate_expression(first_arg, ctx)?;
                            if let Value::String(addr) = address {
                                return Ok(Value::Object(Some(Box::new(Value::String(format!("Range:{}", addr))))));
                            }
                        }
                    }
                }
                // Handle Worksheets("Sheet1").Range("A1")
                if let Expression::FunctionCall { function: inner_fn, args: inner_args } = &**obj {
                    if let Expression::Identifier(fn_name) = &**inner_fn {
                        if fn_name.eq_ignore_ascii_case("Worksheets") && method_name.eq_ignore_ascii_case("Range") {
                            // Worksheets("Sheet1").Range("A1")
                            if let Some(first_arg) = args.first() {
                                let address = evaluate_expression(first_arg, ctx)?;
                                if let Value::String(addr) = address {
                                    return Ok(Value::Object(Some(Box::new(Value::String(format!("Range:{}", addr))))));
                                }
                            }
                        }
                    }
                }
            }

            // Only simple identifier calls supported for now
            let name = if let Expression::Identifier(n) = &**function {
                n
            } else {
                bail!("Only simple identifier calls supported for now")
            };
             // Try builtin functions first
            if let Ok(Some(val)) = crate::interpreter::builtins::functions::handle_builtin_call(name, args, ctx) {
                return Ok(val);
            }
        
            // Built-in: Format(value, pattern)
            if name.eq_ignore_ascii_case("Format") {
                if args.len() != 2 {
                    bail!("Format expects 2 arguments");
                }
                let v = evaluate_expression(&args[0], ctx)?;
                let p = evaluate_expression(&args[1], ctx)?;
                let pat = match p {
                    Value::String(s) => s,
                    other => bail!("Format second argument must be a string, got {:?}", other),
                };
        
                // VBA→chrono pattern adapter
                let mut chrono_pat = pat.clone();
                chrono_pat = chrono_pat.replace("yyyy", "%Y");
                chrono_pat = chrono_pat.replace("mm", "%m");
                chrono_pat = chrono_pat.replace("dd", "%d");
                chrono_pat = chrono_pat.replace("hh", "%H");
                chrono_pat = chrono_pat.replace("ss", "%S");
        
                let s = match v {
                    Value::Date(d) => {
                        let dt = d.and_hms_opt(0, 0, 0).expect("valid midnight");
                        dt.format(&chrono_pat).to_string()
                    }
                    _ => crate::interpreter::coerce::to_string(&v),
                };
                return Ok(Value::String(s));
            }
            if let Expression::Identifier(fn_name) = &**function {
                if fn_name.eq_ignore_ascii_case("Range") {
                    if let Some(first_arg) = args.first() {
                        let address = evaluate_expression(first_arg, ctx)?;
                        if let Value::String(addr) = address {
                            // Range("A1") returns an object reference to the range
                            // We create a special string identifier for the range
                            return Ok(Value::Object(Some(Box::new(Value::String(format!("Range:{}", addr))))));
                        }
                    }
                    bail!("Range() requires a string address argument");
                }
            }
        
            // Try user-defined functions
            if let Some((params, body)) = ctx.subs.get(name).cloned() {
                // Evaluate arguments
                let mut arg_vals = Vec::with_capacity(args.len());
                for a in args.iter() {
                    arg_vals.push(evaluate_expression(a, ctx)?);
                }
                
                // Push a new scope for the function
                ctx.push_scope(name.clone(), crate::context::ScopeKind::Function);
                
                // Bind parameters
                for (param, val) in params.iter().zip(arg_vals.into_iter()) {
                    ctx.declare_variable(&param.name);
                    ctx.declare_local(param.name.clone(), val);
                }
                
                // Initialize the function return variable (FunctionName = ...)
                // In VBA, the function name acts as the return variable
                ctx.declare_variable(name);
                ctx.declare_local(name.clone(), Value::Empty);
                
                // Execute function body
                crate::interpreter::statements::execute_statement_list(&body, ctx);
                
                // Get the return value (the value assigned to the function name)
                let return_value = ctx.get_var(name).unwrap_or(Value::Empty);
                
                // Pop scope
                ctx.pop_scope();
                
                return Ok(return_value);
            }
        
            // Unknown functions -> for now, just return 0
            let _evaluated: Vec<Value> =
                args.iter().map(|a| evaluate_expression(a, ctx)).collect::<Result<_>>()?;
            Ok(Value::Integer(0))
        }        

        // ——— Property Access: Handle enum member access and user types
        PropertyAccess { obj, property } => {
            
        
            // 4) Handle special-case VBA Err object properties
            if let Expression::Identifier(name) = &**obj {
                if name.eq_ignore_ascii_case("Err") {
                    // eprintln!("   → Detected Err object, property={}", property);
                    match property.to_ascii_lowercase().as_str() {
                        "number" => {
                            let n = ctx.err.as_ref().map(|e| e.number).unwrap_or(0);
                            // eprintln!("   → Returning Err.Number = {}", n);
                            return Ok(Value::Integer(n.into()));
                        }
                        "description" => {
                            let d = ctx.err.as_ref()
                                .map(|e| e.description.clone())
                                .unwrap_or_default();
                            // eprintln!("   → Returning Err.Description = {}", d);
                            return Ok(Value::String(d));
                        }
                        "clear" => {
                            // VBA Err.Clear is a subroutine (no return)
                            // eprintln!("   → Calling Err.Clear()");
                            ctx.err = None;
                            ctx.resume_valid = false;
                            return Ok(Value::Integer(0));
                        }
                        "source" => {
                            let s = ctx.err.as_ref()
                                .map(|e| e.source.clone())
                                .unwrap_or_default();
                            // eprintln!("   → Returning Err.Source = {}", s);
                            return Ok(Value::String(s));
                        }
                        _ => bail!("Unknown Err property: {}", property),
                    }
                }
            }
            // ✨ Special handling for ActiveSheet.property, ActiveWorkbook.property, Application.property
            if let Expression::Identifier(obj_name) = &**obj {
                if obj_name.eq_ignore_ascii_case("ActiveSheet") {
                    // Route to worksheet properties
                    match crate::host::excel::properties::get_property("worksheet", "", property, ctx) {
                        Ok(value) => return Ok(value),
                        Err(_) => {}
                    }
                } else if obj_name.eq_ignore_ascii_case("ActiveWorkbook") {
                    // Route to workbook properties
                    match crate::host::excel::properties::get_property("workbook", "", property, ctx) {
                        Ok(value) => return Ok(value),
                        Err(_) => {}
                    }
                } else if obj_name.eq_ignore_ascii_case("Application") {
                    // Route to application properties
                    match crate::host::excel::properties::get_property("application", "", property, ctx) {
                        Ok(value) => return Ok(value),
                        Err(_) => {}
                    }
                }
                
                // Check if it's a registered COM object
                if ctx.com_registry.get_global(obj_name).is_some() {
                    return crate::host::dispatch_com_call(
                        obj_name,
                        property,
                        None,  // No args for property get
                        false, // Not a set
                        ctx,
                    );
                }
                // Check if it's an enum
                if let Some(value) = ctx.get_enum_value(obj_name, property) {
                    return Ok(Value::Integer(value));
                }
            }
            
            // ✨ Special handling for Range().property_name, Range().method_name,
            //    Worksheets().property_name, and also ActiveSheet.Range("A1").property_name
            if let Expression::FunctionCall { function, args } = &**obj {
                // Case 1: Range("A1").Value or Range("B" & i).Value (simple function call)
                if let Expression::Identifier(fn_name) = &**function {
                    if fn_name.eq_ignore_ascii_case("Range") {
                        // Evaluate the argument (supports both literals and expressions like "B" & i)
                        if let Some(arg) = args.first() {
                            let arg_value = evaluate_expression(arg, ctx)?;
                            let address = match arg_value {
                                Value::String(s) => s,
                                other => format!("{:?}", other),
                            };
                            match crate::host::excel::properties::get_property("range", &address, property, ctx) {
                                Ok(value) => return Ok(value),
                                Err(_) => {
                                    return crate::host::excel::methods::call_method("range", &address, property, &[]);
                                }
                            }
                        }
                    }
                    // Case 1b: Worksheets("Sheet1").Name (Worksheets function call)
                    else if fn_name.eq_ignore_ascii_case("Worksheets") {
                        if let Some(Expression::String(sheet_name)) = args.first() {
                            // Format as "name:workbook_id:index" - for Worksheets(), we don't have workbook_id yet
                            // but we can pass just the name and let the handler use empty workbook_id
                            let data = format!("{}::", sheet_name);
                            match crate::host::excel::properties::get_property("worksheet", &data, property, ctx) {
                                Ok(value) => return Ok(value),
                                Err(_) => {
                                    return crate::host::excel::methods::call_method("worksheet", &data, property, &[]);
                                }
                            }
                        }
                    }
                }
                // Case 2: ActiveSheet.Range("A1").Value or ActiveSheet.Range("B" & i).Value (method call on object property)
                else if let Expression::PropertyAccess { obj: _obj_inner, property: inner_prop } = &**function {
                    if inner_prop.eq_ignore_ascii_case("Range") {
                        // Evaluate the argument (supports both literals and expressions like "B" & i)
                        if let Some(arg) = args.first() {
                            let arg_value = evaluate_expression(arg, ctx)?;
                            let address = match arg_value {
                                Value::String(s) => s,
                                other => format!("{:?}", other),
                            };
                            match crate::host::excel::properties::get_property("range", &address, property, ctx) {
                                Ok(value) => return Ok(value),
                                Err(_) => {
                                    return crate::host::excel::methods::call_method("range", &address, property, &[]);
                                }
                            }
                        }
                    }
                }
            }
        

            // 1) Evaluate the object expression first
            let object_val = evaluate_expression(obj, ctx)?;
        
            // 2) Handle user-defined types (Type ... End Type)
            if let Value::UserType { fields, type_name } = &object_val {
                if let Some(val) = fields.get(property) {
                    return Ok(val.clone());
                } else {
                    bail!("Field '{}' not found on type '{}'", property, type_name);
                }
            }
            
            // 2b) Handle object references (Range, Worksheet, etc.)
            if let Value::Object(Some(inner)) = &object_val {
                if let Value::String(obj_ref) = &**inner {
                    // Handle Range:address objects
                    if obj_ref.starts_with("Range:") {
                        let address = &obj_ref[6..]; // Skip "Range:" prefix
                        match crate::host::excel::properties::get_property("range", address, property, ctx) {
                            Ok(value) => return Ok(value),
                            Err(_) => {
                                return crate::host::excel::methods::call_method("range", address, property, &[]);
                            }
                        }
                    }
                }
            }
        
            // 3) Handle enum member access (EnumName.Member)
            if let Expression::Identifier(enum_name) = &**obj {
                if let Some(value) = ctx.get_enum_value(enum_name, property) {
                    return Ok(Value::Integer(value));
                }
            }
        
            // 5) Fallback: if we reach here, property access type was unsupported
            match object_val {
                Value::String(_) | Value::Integer(_) | Value::Boolean(_) => {
                    bail!(
                        "Property access '{:?}.{}' not supported for {:?}",
                        obj, property, object_val
                    )
                }
                other => bail!("Cannot access property '{}' on {:?}", property, other),
            }
        }

        // ——— With Member Access: .Property (within With blocks)
        WithMemberAccess { property } => {
            // Get the current With object from the stack
            if let Some(with_obj) = ctx.with_stack.last().cloned() {
                // Now we need to access the property on the with_obj
                // For Range objects, we need to extract the address and call the property getter
                match &with_obj {
                    Value::Object(Some(inner)) => {
                        if let Value::String(obj_str) = inner.as_ref() {
                            // Check if this is a Range reference
                            if obj_str.to_lowercase().starts_with("range:") {
                                let address = obj_str.strip_prefix("range:").unwrap_or(obj_str);
                                match crate::host::excel::properties::get_property("range", address, property, ctx) {
                                    Ok(value) => return Ok(value),
                                    Err(e) => bail!("Error getting property .{}: {}", property, e),
                                }
                            }
                        }
                        // Try to get field from the object
                        if let Some(val) = inner.get_field(property) {
                            return Ok(val.clone());
                        }
                        bail!("Property '{}' not found on With object", property);
                    }
                    Value::String(obj_str) => {
                        // Check if this is a Range reference stored as string
                        if obj_str.to_lowercase().starts_with("range:") {
                            let address = obj_str.strip_prefix("range:").unwrap_or(obj_str);
                            match crate::host::excel::properties::get_property("range", address, property, ctx) {
                                Ok(value) => return Ok(value),
                                Err(e) => bail!("Error getting property .{}: {}", property, e),
                            }
                        }
                        bail!("Cannot access property '{}' on string value", property);
                    }
                    other => {
                        // Try to get field from the value
                        if let Some(val) = other.get_field(property) {
                            return Ok(val.clone());
                        }
                        bail!("Cannot access property '{}' on {:?}", property, other);
                    }
                }
            } else {
                bail!("'.{}' used outside of With block", property);
            }
        }

        // ——— With Method Call: .Method(args) (within With blocks)
        WithMethodCall { method, args } => {
            // Get the current With object from the stack
            if let Some(with_obj) = ctx.with_stack.last().cloned() {
                // Evaluate method arguments
                let mut evaluated_args = Vec::new();
                for arg in args {
                    evaluated_args.push(evaluate_expression(arg, ctx)?);
                }
                
                // The With object should be a Worksheet, so .Range("A1") means calling Range on that sheet
                match &with_obj {
                    Value::Object(Some(inner)) => {
                        if let Value::String(obj_str) = inner.as_ref() {
                            // Check if this is a Worksheet reference
                            if obj_str.to_lowercase().starts_with("worksheet:") {
                                let sheet_name = obj_str.strip_prefix("worksheet:").unwrap_or(obj_str);
                                
                                // If method is "Range", we need to return a Range object for that sheet
                                if method.eq_ignore_ascii_case("Range") {
                                    if let Some(Value::String(addr)) = evaluated_args.first() {
                                        // Return a Range reference that includes the sheet context
                                        return Ok(Value::Object(Some(Box::new(Value::String(
                                            format!("range:{}!{}", sheet_name, addr)
                                        )))));
                                    }
                                }
                            }
                        }
                        // Generic method call on object
                        bail!("Method '.{}' not supported on With object", method);
                    }
                    _ => {
                        bail!("Cannot call method '.{}' on {:?}", method, with_obj);
                    }
                }
            } else {
                bail!("'.{}()' used outside of With block", method);
            }
        }
    }
}