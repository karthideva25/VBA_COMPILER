//! VBA Interaction Functions
//! 
//! This module contains all VBA interaction and control flow functions including:
//! - IIf — Inline If
//! - Choose, Switch
//! - MsgBox, InputBox (stub implementations)
//! - Shell, Beep, DoEvents
//! - Environ, CurDir, Dir, Command
//! - AppActivate, SendKeys, CreateObject, GetObject

use anyhow::Result;
use crate::ast::Expression;
use crate::context::{Context, Value};
use crate::interpreter::evaluate_expression;
use super::common::{get_optional_int, get_optional_string, value_to_string};

/// Handle interaction-related builtin function calls
pub(crate) fn handle_interaction_function(function: &str, args: &[Expression], ctx: &mut Context) -> Result<Option<Value>> {
    match function {
        // ============================================================
        // CONTROL FLOW FUNCTIONS
        // ============================================================

        // IIF — Returns one of two parts, depending on the evaluation of an expression
        // IIf(expr, truepart, falsepart)
        // IMPORTANT: Unlike If...Then, IIf evaluates BOTH truepart and falsepart
        "iif" => {
            if args.len() < 3 {
                anyhow::bail!("IIf requires 3 arguments: IIf(expr, truepart, falsepart)");
            }
            let condition = evaluate_expression(&args[0], ctx)?;
            let cond_bool = match condition {
                Value::Boolean(b) => b,
                Value::Integer(n) => n != 0,
                Value::Long(n) => n != 0,
                Value::LongLong(n) => n != 0,
                Value::Double(n) => n != 0.0,
                Value::Single(n) => n != 0.0,
                Value::String(s) => !s.is_empty(),
                Value::Empty | Value::Null => false,
                Value::Object(None) => false,  // Nothing is false
                _ => true
            };
            
            // Note: In VBA, IIf evaluates BOTH branches, we return the appropriate one
            // This is different from If...Then which short-circuits
            if cond_bool {
                Ok(Some(evaluate_expression(&args[1], ctx)?))
            } else {
                Ok(Some(evaluate_expression(&args[2], ctx)?))
            }
        }

        // CHOOSE — Selects and returns a value from a list of arguments
        // Choose(index, choice1, [choice2], [choice3], ...)
        // Index is 1-based
        "choose" => {
            if args.is_empty() {
                anyhow::bail!("Choose requires at least 1 argument");
            }
            let index_val = evaluate_expression(&args[0], ctx)?;
            let index = match index_val {
                Value::Integer(n) => n,
                Value::Long(n) => n as i64,
                Value::LongLong(n) => n,
                Value::Double(n) => n.round() as i64,  // VBA rounds to nearest
                Value::Single(n) => n.round() as i64,
                Value::Currency(n) => n.round() as i64,  // Currency also rounds
                Value::Decimal(n) => n.round() as i64,   // Decimal also rounds
                Value::String(s) => s.parse::<i64>().unwrap_or(0),
                Value::Boolean(b) => if b { -1 } else { 0 },  // True = -1 in VBA
                Value::Empty => 0,
                Value::Null => return Ok(Some(Value::Null)),  // Null index returns Null
                _ => 0
            };
            
            // Choose is 1-based, index 0 or negative returns Null
            if index < 1 || index as usize > args.len() - 1 {
                Ok(Some(Value::Null))
            } else {
                Ok(Some(evaluate_expression(&args[index as usize], ctx)?))
            }
        }

        // SWITCH — Evaluates a list of expressions and returns a value
        // Switch(expr1, value1, [expr2, value2], ...)
        // Returns value for first True expression, Null if none match
        "switch" => {
            // Switch takes pairs of (condition, value)
            if args.len() < 2 {
                anyhow::bail!("Switch requires at least one pair: Switch(expr1, value1, ...)");
            }
            if args.len() % 2 != 0 {
                anyhow::bail!("Switch requires pairs of expressions: Switch(expr1, value1, expr2, value2, ...)");
            }
            
            for chunk in args.chunks(2) {
                let condition = evaluate_expression(&chunk[0], ctx)?;
                let cond_bool = match condition {
                    Value::Boolean(b) => b,
                    Value::Integer(n) => n != 0,
                    Value::Long(n) => n != 0,
                    Value::LongLong(n) => n != 0,
                    Value::Double(n) => n != 0.0,
                    Value::Single(n) => n != 0.0,
                    Value::String(s) => !s.is_empty(),  // Non-empty string is truthy
                    Value::Empty | Value::Null => false,
                    _ => false
                };
                
                if cond_bool {
                    return Ok(Some(evaluate_expression(&chunk[1], ctx)?));
                }
            }
            
            // No match found - return Null
            Ok(Some(Value::Null))
        }

        // ============================================================
        // MESSAGE/INPUT FUNCTIONS
        // ============================================================

        // MSGBOX — Displays a message in a dialog box
        // MsgBox(Prompt, [Buttons], [Title], [HelpFile], [Context])
        // Buttons constants:
        //   vbOKOnly = 0, vbOKCancel = 1, vbAbortRetryIgnore = 2
        //   vbYesNoCancel = 3, vbYesNo = 4, vbRetryCancel = 5
        // Return values:
        //   vbOK = 1, vbCancel = 2, vbAbort = 3, vbRetry = 4
        //   vbIgnore = 5, vbYes = 6, vbNo = 7
        "msgbox" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(1))); // vbOK
            }
            
            let message = evaluate_expression(&args[0], ctx)?;
            let message_str = value_to_string(&message);
            
            // Log to context output for testing
            ctx.log(&format!("MsgBox: {}", message_str));
            
            // Get buttons parameter (default 0 = vbOKOnly)
            let buttons = get_optional_int(args, 1, 0, ctx)?;
            
            // Return appropriate default button based on button style
            // Lower 4 bits determine button configuration
            let button_type = buttons & 0x0F;
            let default_return = match button_type {
                0 => 1,  // vbOKOnly -> vbOK (1)
                1 => 1,  // vbOKCancel -> vbOK (1)
                2 => 3,  // vbAbortRetryIgnore -> vbAbort (3)
                3 => 6,  // vbYesNoCancel -> vbYes (6)
                4 => 6,  // vbYesNo -> vbYes (6)
                5 => 4,  // vbRetryCancel -> vbRetry (4)
                _ => 1,  // Default to vbOK
            };
            
            Ok(Some(Value::Integer(default_return)))
        }

        // INPUTBOX — Displays a prompt in a dialog box, waits for user input
        // InputBox(Prompt, [Title], [Default], [XPos], [YPos], [HelpFile], [Context])
        // In non-interactive mode:
        //   1. Returns mock value if set in context
        //   2. Returns Default parameter if provided
        //   3. Returns empty string otherwise
        "inputbox" => {
            // Check if there's a mock input value set in context
            if let Some(mock_value) = ctx.get_var("__INPUT_MOCK__") {
                return Ok(Some(mock_value.clone()));
            }
            
            // Get default value (3rd parameter, index 2)
            let default_value = get_optional_string(args, 2, "", ctx)?;
            
            Ok(Some(Value::String(default_value)))
        }

        // ============================================================
        // SYSTEM INTERACTION
        // ============================================================

        // BEEP — Sounds a tone through the computer's speaker
        "beep" => {
            // No-op in this implementation (no sound)
            ctx.log("Beep");
            Ok(Some(Value::Empty))
        }

        // SHELL — Runs an executable program
        // Shell(PathName, [WindowStyle])
        // SECURITY: Returns 0 (disabled) - executing arbitrary commands is dangerous
        "shell" => {
            // Log for debugging/testing
            if !args.is_empty() {
                let path = evaluate_expression(&args[0], ctx)?;
                ctx.log(&format!("Shell (blocked): {}", value_to_string(&path)));
            }
            // Return 0 (no process ID) for security
            Ok(Some(Value::Double(0.0)))
        }

        // DOEVENTS — Yields execution so the OS can process other events
        // Returns number of open forms (0 in our implementation)
        "doevents" => {
            // No-op in this implementation
            Ok(Some(Value::Integer(0)))
        }

        // ENVIRON — Returns the string associated with an OS environment variable
        // Environ(EnvString) or Environ(Number)
        "environ" | "environ$" => {
            if args.is_empty() {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let result = match val {
                Value::String(name) => {
                    // Look up by name
                    std::env::var(&name).unwrap_or_default()
                }
                Value::Integer(n) => {
                    // Look up by index (1-based)
                    if n < 1 {
                        String::new()
                    } else {
                        std::env::vars()
                            .nth((n - 1) as usize)
                            .map(|(k, v)| format!("{}={}", k, v))
                            .unwrap_or_default()
                    }
                }
                Value::Long(n) => {
                    // Look up by index (1-based)
                    if n < 1 {
                        String::new()
                    } else {
                        std::env::vars()
                            .nth((n - 1) as usize)
                            .map(|(k, v)| format!("{}={}", k, v))
                            .unwrap_or_default()
                    }
                }
                Value::Double(n) => {
                    let n = n as i64;
                    if n < 1 {
                        String::new()
                    } else {
                        std::env::vars()
                            .nth((n - 1) as usize)
                            .map(|(k, v)| format!("{}={}", k, v))
                            .unwrap_or_default()
                    }
                }
                _ => String::new()
            };
            Ok(Some(Value::String(result)))
        }

        // COMMAND — Returns the argument portion of the command line
        // Command$ is the string version
        "command" | "command$" => {
            // Get command line arguments (skip program name)
            let args: Vec<String> = std::env::args().skip(1).collect();
            Ok(Some(Value::String(args.join(" "))))
        }

        // CURDIR — Returns the current path
        // CurDir([drive]) - drive parameter is ignored in modern systems
        "curdir" | "curdir$" => {
            let path = std::env::current_dir()
                .map(|p| p.to_string_lossy().to_string())
                .unwrap_or_default();
            Ok(Some(Value::String(path)))
        }

        // DIR — Returns a file/directory name matching a pattern
        // Dir([PathName], [Attributes])
        // First call with pattern returns first match, subsequent calls without args return next
        "dir" | "dir$" => {
            // Simplified implementation - returns empty string
            // Full implementation would need stateful iteration
            if !args.is_empty() {
                let _pattern = evaluate_expression(&args[0], ctx)?;
                // TODO: Implement directory listing with pattern matching
            }
            Ok(Some(Value::String(String::new())))
        }

        // ============================================================
        // APPLICATION CONTROL (STUBS)
        // ============================================================

        // APPACTIVATE — Activates an application window
        // AppActivate(Title, [Wait])
        "appactivate" => {
            // No-op - can't activate windows in this context
            if !args.is_empty() {
                let title = evaluate_expression(&args[0], ctx)?;
                ctx.log(&format!("AppActivate (stub): {}", value_to_string(&title)));
            }
            Ok(Some(Value::Empty))
        }

        // SENDKEYS — Sends keystrokes to the active window
        // SendKeys(String, [Wait])
        // SECURITY: Disabled - sending keystrokes can be dangerous
        "sendkeys" => {
            if !args.is_empty() {
                let keys = evaluate_expression(&args[0], ctx)?;
                ctx.log(&format!("SendKeys (blocked): {}", value_to_string(&keys)));
            }
            Ok(Some(Value::Empty))
        }

        // CREATEOBJECT — Creates an OLE Automation object
        // CreateObject(Class, [ServerName])
        "createobject" => {
            if args.is_empty() {
                anyhow::bail!("CreateObject requires a class name");
            }
            let class_name = evaluate_expression(&args[0], ctx)?;
            let class_str = value_to_string(&class_name);
            ctx.log(&format!("CreateObject (stub): {}", class_str));
            
            // Return a stub object
            Ok(Some(Value::Object(Some(Box::new(Value::String(class_str))))))
        }

        // GETOBJECT — Returns a reference to an object provided by an OLE server
        // GetObject([PathName], [Class])
        "getobject" => {
            let path = get_optional_string(args, 0, "", ctx)?;
            let class = get_optional_string(args, 1, "", ctx)?;
            ctx.log(&format!("GetObject (stub): path={}, class={}", path, class));
            
            // Return a stub object or Nothing
            if path.is_empty() && class.is_empty() {
                Ok(Some(Value::Object(None)))  // Nothing
            } else {
                Ok(Some(Value::Object(Some(Box::new(Value::String(
                    if !class.is_empty() { class } else { path }
                ))))))
            }
        }

        _ => Ok(None)
    }
}
