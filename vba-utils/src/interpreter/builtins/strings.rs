//! VBA String Functions
//! 
//! This module contains all VBA string manipulation functions including:
//! - Len, LenB
//! - Left, Left$, LeftB, LeftB$
//! - Right, Right$, RightB, RightB$
//! - Mid, Mid$, MidB, MidB$
//! - UCase, UCase$, LCase, LCase$
//! - Trim, Trim$, LTrim, LTrim$, RTrim, RTrim$
//! - InStr, InStrB, InStrRev
//! - Replace, StrReverse
//! - Asc, AscB, AscW
//! - Chr, Chr$, ChrB, ChrB$, ChrW, ChrW$
//! - Space, Space$, String, String$
//! - StrComp, StrConv
//! - Format, Format$, FormatCurrency, FormatNumber, FormatPercent

use anyhow::Result;
use crate::ast::Expression;
use crate::context::{Context, Value};
use crate::interpreter::evaluate_expression;
use super::common::value_to_string;

/// Handle string-related builtin function calls
pub(crate) fn handle_string_function(function: &str, args: &[Expression], ctx: &mut Context) -> Result<Option<Value>> {
    match function {
        // ============================================================
        // BASIC STRING FUNCTIONS
        // ============================================================
        
        // LEN — returns length of string
        "len" => {
            if args.len() != 1 {
                ctx.log("*** Error: Len() expects 1 argument");
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => Ok(Some(Value::Integer(s.len() as i64))),
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // LENB — LenB(string) returns byte length (UTF-16 in VBA)
        "lenb" => {
            if args.len() != 1 {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => Ok(Some(Value::Integer((s.len() * 2) as i64))), // UTF-16 bytes
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // ============================================================
        // SUBSTRING FUNCTIONS
        // ============================================================

        // MID — Mid(string, start, [length])
        "mid" | "mid$" => {
            if args.len() < 2 || args.len() > 3 {
                return Ok(Some(Value::String(String::new())));
            }
            let string_val = evaluate_expression(&args[0], ctx)?;
            let start_val = evaluate_expression(&args[1], ctx)?;
            let s = match string_val { Value::String(s) => s, _ => return Ok(Some(Value::String(String::new()))) };
            let start = match start_val { Value::Integer(i) => (i - 1).max(0) as usize, _ => return Ok(Some(Value::String(String::new()))) };
            
            if args.len() == 3 {
                let len_val = evaluate_expression(&args[2], ctx)?;
                let len = match len_val { Value::Integer(i) => i.max(0) as usize, _ => return Ok(Some(Value::String(String::new()))) };
                let result: String = s.chars().skip(start).take(len).collect();
                Ok(Some(Value::String(result)))
            } else {
                let result: String = s.chars().skip(start).collect();
                Ok(Some(Value::String(result)))
            }
        }

        // MIDB — MidB(string, start, [length]) - byte-based
        "midb" | "midb$" => {
            if args.len() < 2 || args.len() > 3 {
                return Ok(Some(Value::String(String::new())));
            }
            let string_val = evaluate_expression(&args[0], ctx)?;
            let start_val = evaluate_expression(&args[1], ctx)?;
            let s = match string_val { Value::String(s) => s, _ => return Ok(Some(Value::String(String::new()))) };
            let start = match start_val { Value::Integer(i) => ((i - 1) / 2).max(0) as usize, _ => return Ok(Some(Value::String(String::new()))) };
            
            if args.len() == 3 {
                let len_val = evaluate_expression(&args[2], ctx)?;
                let len = match len_val { Value::Integer(i) => (i / 2).max(0) as usize, _ => return Ok(Some(Value::String(String::new()))) };
                let result: String = s.chars().skip(start).take(len).collect();
                Ok(Some(Value::String(result)))
            } else {
                let result: String = s.chars().skip(start).collect();
                Ok(Some(Value::String(result)))
            }
        }

        // LEFT — Left(string, length)
        "left" | "left$" => {
            if args.len() != 2 {
                return Ok(Some(Value::String(String::new())));
            }
            let string_val = evaluate_expression(&args[0], ctx)?;
            let length_val = evaluate_expression(&args[1], ctx)?;
            match (string_val, length_val) {
                (Value::String(s), Value::Integer(len)) => {
                    let len = len.max(0) as usize;
                    let result: String = s.chars().take(len).collect();
                    Ok(Some(Value::String(result)))
                }
                _ => Ok(Some(Value::String(String::new())))
            }
        }

        // LEFTB — LeftB(string, length) - byte-based
        "leftb" | "leftb$" => {
            if args.len() != 2 {
                return Ok(Some(Value::String(String::new())));
            }
            let string_val = evaluate_expression(&args[0], ctx)?;
            let length_val = evaluate_expression(&args[1], ctx)?;
            match (string_val, length_val) {
                (Value::String(s), Value::Integer(len)) => {
                    let byte_len = (len / 2).max(0) as usize;
                    let result: String = s.chars().take(byte_len).collect();
                    Ok(Some(Value::String(result)))
                }
                _ => Ok(Some(Value::String(String::new())))
            }
        }

        // RIGHT — Right(string, length)
        "right" | "right$" => {
            if args.len() != 2 {
                return Ok(Some(Value::String(String::new())));
            }
            let string_val = evaluate_expression(&args[0], ctx)?;
            let length_val = evaluate_expression(&args[1], ctx)?;
            match (string_val, length_val) {
                (Value::String(s), Value::Integer(len)) => {
                    let len = len.max(0) as usize;
                    let char_count = s.chars().count();
                    let skip = char_count.saturating_sub(len);
                    let result: String = s.chars().skip(skip).collect();
                    Ok(Some(Value::String(result)))
                }
                _ => Ok(Some(Value::String(String::new())))
            }
        }

        // RIGHTB — RightB(string, length) - byte-based
        "rightb" | "rightb$" => {
            if args.len() != 2 {
                return Ok(Some(Value::String(String::new())));
            }
            let string_val = evaluate_expression(&args[0], ctx)?;
            let length_val = evaluate_expression(&args[1], ctx)?;
            match (string_val, length_val) {
                (Value::String(s), Value::Integer(len)) => {
                    let byte_len = (len / 2).max(0) as usize;
                    let char_count = s.chars().count();
                    let skip = char_count.saturating_sub(byte_len);
                    let result: String = s.chars().skip(skip).collect();
                    Ok(Some(Value::String(result)))
                }
                _ => Ok(Some(Value::String(String::new())))
            }
        }

        // ============================================================
        // CASE CONVERSION
        // ============================================================

        // UCASE — UCase(string)
        "ucase" | "ucase$" => {
            if args.len() != 1 {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => Ok(Some(Value::String(s.to_uppercase()))),
                _ => Ok(Some(Value::String(String::new())))
            }
        }

        // LCASE — LCase(string)
        "lcase" | "lcase$" => {
            if args.len() != 1 {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => Ok(Some(Value::String(s.to_lowercase()))),
                _ => Ok(Some(Value::String(String::new())))
            }
        }

        // ============================================================
        // TRIM FUNCTIONS
        // ============================================================

        // TRIM — Trim(string)
        "trim" | "trim$" => {
            if args.len() != 1 {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => Ok(Some(Value::String(s.trim().to_string()))),
                _ => Ok(Some(Value::String(String::new())))
            }
        }

        // LTRIM — LTrim(string)
        "ltrim" | "ltrim$" => {
            if args.len() != 1 {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => Ok(Some(Value::String(s.trim_start().to_string()))),
                _ => Ok(Some(Value::String(String::new())))
            }
        }

        // RTRIM — RTrim(string)
        "rtrim" | "rtrim$" => {
            if args.len() != 1 {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => Ok(Some(Value::String(s.trim_end().to_string()))),
                _ => Ok(Some(Value::String(String::new())))
            }
        }

        // ============================================================
        // SEARCH FUNCTIONS
        // ============================================================

        // INSTR — InStr([start,] string1, string2, [compare])
        // VBA allows two calling conventions:
        // 1. InStr(string1, string2) - start defaults to 1
        // 2. InStr(start, string1, string2) - explicit start position
        // 3. InStr(start, string1, string2, compare) - with compare mode
        // compare: vbBinaryCompare=0 (default), vbTextCompare=1
        "instr" => {
            if args.len() < 2 || args.len() > 4 {
                return Ok(Some(Value::Integer(0)));
            }

            // Determine calling convention based on first argument type and arg count
            let (start, str1, str2, compare) = if args.len() == 2 {
                // InStr(string1, string2) - 2 arg form
                let s1 = super::common::get_required_string(args, 0, ctx)?;
                let s2 = super::common::get_required_string(args, 1, ctx)?;
                (1i64, s1, s2, 0i64)
            } else if args.len() >= 3 {
                // Check if first arg is numeric (start position) or string
                let first_val = evaluate_expression(&args[0], ctx)?;
                match first_val {
                    Value::Integer(_) | Value::Long(_) | Value::Double(_) => {
                        // InStr(start, string1, string2, [compare])
                        let start = super::common::get_required_int(args, 0, ctx)?;
                        let s1 = super::common::get_required_string(args, 1, ctx)?;
                        let s2 = super::common::get_required_string(args, 2, ctx)?;
                        let cmp = super::common::get_optional_int(args, 3, 0, ctx)?;
                        (start, s1, s2, cmp)
                    }
                    Value::String(s1) => {
                        // InStr(string1, string2, [compare]) - rare but valid
                        let s2 = super::common::get_required_string(args, 1, ctx)?;
                        let cmp = super::common::get_optional_int(args, 2, 0, ctx)?;
                        (1, s1, s2, cmp)
                    }
                    _ => return Ok(Some(Value::Integer(0)))
                }
            } else {
                return Ok(Some(Value::Integer(0)));
            };

            // Handle special cases
            if str2.is_empty() {
                return Ok(Some(Value::Integer(start))); // VBA returns start for empty search string
            }
            if str1.is_empty() || start < 1 {
                return Ok(Some(Value::Integer(0)));
            }

            let start_idx = ((start - 1).max(0) as usize).min(str1.len());
            
            // Perform search based on compare mode
            let result = if compare == 1 {
                // Case-insensitive search
                let str1_lower = str1.to_lowercase();
                let str2_lower = str2.to_lowercase();
                str1_lower[start_idx..].find(&str2_lower)
            } else {
                // Case-sensitive search (default)
                str1[start_idx..].find(&str2)
            };
            
            match result {
                Some(pos) => Ok(Some(Value::Integer((start_idx + pos + 1) as i64))),
                None => Ok(Some(Value::Integer(0)))
            }
        }

        // INSTRB — InStrB([start,] string1, string2)
        "instrb" => {
            if args.len() < 2 {
                return Ok(Some(Value::Integer(0)));
            }
            let str1 = super::common::get_required_string(args, 0, ctx)?;
            let str2 = super::common::get_required_string(args, 1, ctx)?;
            
            match str1.find(&str2) {
                Some(pos) => Ok(Some(Value::Integer(((pos + 1) * 2) as i64))),
                None => Ok(Some(Value::Integer(0)))
            }
        }

        // INSTRREV — InStrRev(stringcheck, stringmatch, [start], [compare])
        // start: Position to start searching from (default -1 = end of string)
        // compare: vbBinaryCompare=0 (default), vbTextCompare=1
        "instrrev" => {
            if args.len() < 2 {
                return Ok(Some(Value::Integer(0)));
            }
            
            let str1 = super::common::get_required_string(args, 0, ctx)?;
            let str2 = super::common::get_required_string(args, 1, ctx)?;
            let start = super::common::get_optional_int(args, 2, -1, ctx)?;
            let compare = super::common::get_optional_int(args, 3, 0, ctx)?;
            
            if str2.is_empty() {
                return Ok(Some(Value::Integer(if start < 0 { str1.len() as i64 } else { start })));
            }
            if str1.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            
            // Determine search range
            let search_str = if start < 0 {
                str1.as_str()
            } else {
                let end_idx = (start as usize).min(str1.len());
                &str1[..end_idx]
            };
            
            // Perform reverse search based on compare mode
            let result = if compare == 1 {
                // Case-insensitive search
                let search_lower = search_str.to_lowercase();
                let str2_lower = str2.to_lowercase();
                search_lower.rfind(&str2_lower)
            } else {
                // Case-sensitive search (default)
                search_str.rfind(&str2)
            };
            
            match result {
                Some(pos) => Ok(Some(Value::Integer((pos + 1) as i64))),
                None => Ok(Some(Value::Integer(0)))
            }
        }

        // ============================================================
        // REPLACE AND MANIPULATION
        // ============================================================

        // REPLACE — Replace(expression, find, replace, [start], [count], [compare])
        // start: Position to start in expression (default 1)
        // count: Number of replacements to make (-1 = all, default)
        // compare: vbBinaryCompare=0 (default), vbTextCompare=1
        "replace" => {
            if args.len() < 3 {
                return Ok(Some(Value::String(String::new())));
            }
            
            let expr = super::common::get_required_string(args, 0, ctx)?;
            let find = super::common::get_required_string(args, 1, ctx)?;
            let repl = super::common::get_required_string(args, 2, ctx)?;
            let start = super::common::get_optional_int(args, 3, 1, ctx)? as usize;
            let count = super::common::get_optional_int(args, 4, -1, ctx)?;
            let compare = super::common::get_optional_int(args, 5, 0, ctx)?;
            
            if find.is_empty() {
                return Ok(Some(Value::String(expr)));
            }
            
            // Adjust for 1-based VBA indexing
            let start_idx = start.saturating_sub(1).min(expr.len());
            
            // Split string at start position
            let prefix: String = expr.chars().take(start_idx).collect();
            let work_str: String = expr.chars().skip(start_idx).collect();
            
            // Perform replacement based on compare mode
            let result = if compare == 1 {
                // Case-insensitive replacement
                let find_lower = find.to_lowercase();
                let mut result = String::new();
                let mut remaining = work_str.as_str();
                let mut replacements = 0i64;
                
                while !remaining.is_empty() {
                    if count >= 0 && replacements >= count {
                        result.push_str(remaining);
                        break;
                    }
                    
                    if let Some(pos) = remaining.to_lowercase().find(&find_lower) {
                        result.push_str(&remaining[..pos]);
                        result.push_str(&repl);
                        remaining = &remaining[pos + find.len()..];
                        replacements += 1;
                    } else {
                        result.push_str(remaining);
                        break;
                    }
                }
                result
            } else {
                // Case-sensitive replacement (default)
                if count < 0 {
                    work_str.replace(&find, &repl)
                } else {
                    work_str.replacen(&find, &repl, count as usize)
                }
            };
            
            Ok(Some(Value::String(format!("{}{}", prefix, result))))
        }

        // STRREVERSE — StrReverse(string)
        "strreverse" => {
            if args.len() != 1 {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => Ok(Some(Value::String(s.chars().rev().collect()))),
                _ => Ok(Some(Value::String(String::new())))
            }
        }

        // ============================================================
        // CHARACTER FUNCTIONS
        // ============================================================

        // ASC — Asc(string) returns ASCII code of first character
        "asc" => {
            if args.len() != 1 {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => {
                    if let Some(c) = s.chars().next() {
                        Ok(Some(Value::Integer(c as i64)))
                    } else {
                        Ok(Some(Value::Integer(0)))
                    }
                }
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // ASCB — AscB(string) returns byte value of first character
        "ascb" => {
            if args.len() != 1 {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => {
                    if let Some(b) = s.bytes().next() {
                        Ok(Some(Value::Integer(b as i64)))
                    } else {
                        Ok(Some(Value::Integer(0)))
                    }
                }
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // ASCW — AscW(string) returns Unicode code of first character
        "ascw" => {
            if args.len() != 1 {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::String(s) => {
                    if let Some(c) = s.chars().next() {
                        Ok(Some(Value::Integer(c as u32 as i64)))
                    } else {
                        Ok(Some(Value::Integer(0)))
                    }
                }
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // CHR — Chr(charcode) returns character from ASCII code
        "chr" | "chr$" => {
            if args.len() != 1 {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let code = match val {
                Value::Integer(i) => i,
                Value::Long(l) => l as i64,
                _ => return Ok(Some(Value::String(String::new())))
            };
            if code >= 0 && code <= 255 {
                Ok(Some(Value::String((code as u8 as char).to_string())))
            } else {
                Ok(Some(Value::String(String::new())))
            }
        }

        // CHRB — ChrB(charcode) returns character from byte value
        "chrb" | "chrb$" => {
            if args.len() != 1 {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let code = match val {
                Value::Integer(i) => i,
                Value::Byte(b) => b as i64,
                _ => return Ok(Some(Value::String(String::new())))
            };
            if code >= 0 && code <= 255 {
                Ok(Some(Value::String((code as u8 as char).to_string())))
            } else {
                Ok(Some(Value::String(String::new())))
            }
        }

        // CHRW — ChrW(charcode) returns character from Unicode code
        "chrw" | "chrw$" => {
            if args.len() != 1 {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let code = match val {
                Value::Integer(i) => i as u32,
                Value::Long(l) => l as u32,
                _ => return Ok(Some(Value::String(String::new())))
            };
            if let Some(c) = char::from_u32(code) {
                Ok(Some(Value::String(c.to_string())))
            } else {
                Ok(Some(Value::String(String::new())))
            }
        }

        // ============================================================
        // GENERATION FUNCTIONS
        // ============================================================

        // SPACE — Space(number) returns string of spaces
        "space" | "space$" => {
            if args.len() != 1 {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let count = match val {
                Value::Integer(i) => i.max(0) as usize,
                Value::Long(l) => l.max(0) as usize,
                _ => 0
            };
            Ok(Some(Value::String(" ".repeat(count))))
        }

        // STRING — String(number, character) returns repeated character
        "string" | "string$" => {
            if args.len() != 2 {
                return Ok(Some(Value::String(String::new())));
            }
            let count_val = evaluate_expression(&args[0], ctx)?;
            let char_val = evaluate_expression(&args[1], ctx)?;
            
            let count = match count_val {
                Value::Integer(i) => i.max(0) as usize,
                Value::Long(l) => l.max(0) as usize,
                _ => 0
            };
            
            let ch = match char_val {
                Value::String(s) => s.chars().next().unwrap_or(' '),
                Value::Integer(i) => (i as u8) as char,
                _ => ' '
            };
            
            Ok(Some(Value::String(ch.to_string().repeat(count))))
        }

        // ============================================================
        // COMPARISON AND CONVERSION
        // ============================================================

        // STRCOMP — StrComp(string1, string2, [compare])
        "strcomp" => {
            if args.len() < 2 {
                return Ok(Some(Value::Integer(0)));
            }
            let str1_val = evaluate_expression(&args[0], ctx)?;
            let str2_val = evaluate_expression(&args[1], ctx)?;
            let compare = if args.len() > 2 {
                match evaluate_expression(&args[2], ctx)? {
                    Value::Integer(i) => i,
                    _ => 0
                }
            } else { 0 };
            
            let str1 = match str1_val { Value::String(s) => s, _ => return Ok(Some(Value::Integer(0))) };
            let str2 = match str2_val { Value::String(s) => s, _ => return Ok(Some(Value::Integer(0))) };
            
            let result = if compare == 1 {
                // vbTextCompare - case insensitive
                str1.to_lowercase().cmp(&str2.to_lowercase())
            } else {
                // vbBinaryCompare - case sensitive
                str1.cmp(&str2)
            };
            
            Ok(Some(Value::Integer(match result {
                std::cmp::Ordering::Less => -1,
                std::cmp::Ordering::Equal => 0,
                std::cmp::Ordering::Greater => 1,
            })))
        }

        // STRCONV — StrConv(string, conversion, [localeid])
        "strconv" => {
            if args.len() < 2 {
                return Ok(Some(Value::String(String::new())));
            }
            let str_val = evaluate_expression(&args[0], ctx)?;
            let conv_val = evaluate_expression(&args[1], ctx)?;
            
            let s = match str_val { Value::String(s) => s, _ => return Ok(Some(Value::String(String::new()))) };
            let conv = match conv_val { Value::Integer(i) => i, _ => return Ok(Some(Value::String(s))) };
            
            let result = match conv {
                1 => s.to_uppercase(),  // vbUpperCase
                2 => s.to_lowercase(),  // vbLowerCase
                3 => {
                    // vbProperCase - capitalize first letter of each word
                    s.split_whitespace()
                        .map(|word| {
                            let mut chars: Vec<char> = word.chars().collect();
                            if let Some(first) = chars.first_mut() {
                                *first = first.to_uppercase().next().unwrap_or(*first);
                            }
                            for c in chars.iter_mut().skip(1) {
                                *c = c.to_lowercase().next().unwrap_or(*c);
                            }
                            chars.into_iter().collect::<String>()
                        })
                        .collect::<Vec<_>>()
                        .join(" ")
                }
                _ => s
            };
            Ok(Some(Value::String(result)))
        }

        // ============================================================
        // FORMAT FUNCTIONS
        // ============================================================

        // FORMAT / FORMAT$ — Format(expression, format, [firstdayofweek], [firstweekofyear])
        "format" | "format$" => {
            if args.is_empty() {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let fmt = if args.len() > 1 {
                match evaluate_expression(&args[1], ctx)? {
                    Value::String(s) => s,
                    _ => String::new()
                }
            } else {
                String::new()
            };
            
            let result = match fmt.to_lowercase().as_str() {
                "" | "general" | "standard" => value_to_string(&val),
                "currency" => format_currency(&val),
                "percent" => format_percent(&val, 2),
                "scientific" => format_scientific(&val),
                "yes/no" => format_yes_no(&val),
                "true/false" => format_true_false(&val),
                "on/off" => format_on_off(&val),
                "long date" => format_long_date(&val),
                "short date" => format_short_date(&val),
                "long time" => format_long_time(&val),
                "short time" => format_short_time(&val),
                _ => format_custom(&val, &fmt)
            };
            Ok(Some(Value::String(result)))
        }

        // FORMATCURRENCY — FormatCurrency(expression, [numdigits], ...)
        "formatcurrency" => {
            if args.is_empty() {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let digits = if args.len() > 1 {
                match evaluate_expression(&args[1], ctx)? {
                    Value::Integer(i) => i as usize,
                    _ => 2
                }
            } else { 2 };
            
            let n = value_to_number(&val);
            let formatted = format!("{:.width$}", n.abs(), width = digits);
            let parts: Vec<&str> = formatted.split('.').collect();
            let int_part = parts.get(0).unwrap_or(&"0");
            let dec_part = parts.get(1).unwrap_or(&"00");
            
            let int_with_commas: String = int_part.chars().rev()
                .enumerate()
                .fold(String::new(), |mut acc, (i, c)| {
                    if i > 0 && i % 3 == 0 { acc.push(','); }
                    acc.push(c);
                    acc
                })
                .chars().rev().collect();
            
            if n < 0.0 {
                Ok(Some(Value::String(format!("-${}.{}", int_with_commas, dec_part))))
            } else {
                Ok(Some(Value::String(format!("${}.{}", int_with_commas, dec_part))))
            }
        }

        // FORMATNUMBER — FormatNumber(expression, [numdigits], ...)
        "formatnumber" => {
            if args.is_empty() {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let digits = if args.len() > 1 {
                match evaluate_expression(&args[1], ctx)? {
                    Value::Integer(i) => i as usize,
                    _ => 2
                }
            } else { 2 };
            
            let n = value_to_number(&val);
            Ok(Some(Value::String(format!("{:.width$}", n, width = digits))))
        }

        // FORMATPERCENT — FormatPercent(expression, [numdigits], ...)
        "formatpercent" => {
            if args.is_empty() {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let digits = if args.len() > 1 {
                match evaluate_expression(&args[1], ctx)? {
                    Value::Integer(i) => i as usize,
                    _ => 2
                }
            } else { 2 };
            
            let n = value_to_number(&val);
            Ok(Some(Value::String(format!("{:.width$}%", n * 100.0, width = digits))))
        }

        // WEEKDAYNAME — WeekdayName(weekday, [abbreviate], [firstdayofweek])
        "weekdayname" => {
            if args.is_empty() {
                return Ok(Some(Value::String(String::new())));
            }
            let weekday_val = evaluate_expression(&args[0], ctx)?;
            let abbreviate = if args.len() > 1 {
                match evaluate_expression(&args[1], ctx)? {
                    Value::Boolean(b) => b,
                    _ => false
                }
            } else { false };
            
            let weekday = match weekday_val {
                Value::Integer(i) => i,
                _ => return Ok(Some(Value::String(String::new())))
            };
            
            let names_full = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
            let names_abbrev = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
            
            if weekday >= 1 && weekday <= 7 {
                let name = if abbreviate {
                    names_abbrev[(weekday - 1) as usize]
                } else {
                    names_full[(weekday - 1) as usize]
                };
                Ok(Some(Value::String(name.to_string())))
            } else {
                Ok(Some(Value::String(String::new())))
            }
        }

        _ => Ok(None)
    }
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================

fn value_to_number(val: &Value) -> f64 {
    match val {
        Value::Integer(i) => *i as f64,
        Value::Long(l) => *l as f64,
        Value::Double(d) => *d,
        Value::Single(s) => *s as f64,
        Value::Currency(c) => *c,
        Value::String(s) => s.parse().unwrap_or(0.0),
        _ => 0.0
    }
}

fn format_currency(val: &Value) -> String {
    let n = value_to_number(val);
    let formatted = format!("{:.2}", n.abs());
    let parts: Vec<&str> = formatted.split('.').collect();
    let int_part = parts.get(0).unwrap_or(&"0");
    let dec_part = parts.get(1).unwrap_or(&"00");
    
    let int_with_commas: String = int_part.chars().rev()
        .enumerate()
        .fold(String::new(), |mut acc, (i, c)| {
            if i > 0 && i % 3 == 0 { acc.push(','); }
            acc.push(c);
            acc
        })
        .chars().rev().collect();
    
    if n < 0.0 {
        format!("-${}.{}", int_with_commas, dec_part)
    } else {
        format!("${}.{}", int_with_commas, dec_part)
    }
}

fn format_percent(val: &Value, digits: usize) -> String {
    let n = value_to_number(val);
    format!("{:.width$}%", n * 100.0, width = digits)
}

fn format_scientific(val: &Value) -> String {
    let n = value_to_number(val);
    format!("{:.2E}", n)
}

fn format_yes_no(val: &Value) -> String {
    let b = match val {
        Value::Boolean(b) => *b,
        Value::Integer(i) => *i != 0,
        _ => false
    };
    if b { "Yes".to_string() } else { "No".to_string() }
}

fn format_true_false(val: &Value) -> String {
    let b = match val {
        Value::Boolean(b) => *b,
        Value::Integer(i) => *i != 0,
        _ => false
    };
    if b { "True".to_string() } else { "False".to_string() }
}

fn format_on_off(val: &Value) -> String {
    let b = match val {
        Value::Boolean(b) => *b,
        Value::Integer(i) => *i != 0,
        _ => false
    };
    if b { "On".to_string() } else { "Off".to_string() }
}

fn format_long_date(val: &Value) -> String {
    if let Value::Date(d) = val {
        d.format("%B %d, %Y").to_string()
    } else {
        value_to_string(val)
    }
}

fn format_short_date(val: &Value) -> String {
    if let Value::Date(d) = val {
        d.format("%m/%d/%Y").to_string()
    } else {
        value_to_string(val)
    }
}

fn format_long_time(val: &Value) -> String {
    if let Value::Date(d) = val {
        let dt = d.and_hms_opt(0, 0, 0).unwrap();
        dt.format("%H:%M:%S").to_string()
    } else {
        value_to_string(val)
    }
}

fn format_short_time(val: &Value) -> String {
    if let Value::Date(d) = val {
        let dt = d.and_hms_opt(0, 0, 0).unwrap();
        dt.format("%H:%M").to_string()
    } else {
        value_to_string(val)
    }
}

fn format_custom(val: &Value, fmt: &str) -> String {
    // Extract the datetime to format (either from Date, DateTime, or Time)
    let dt = match val {
        Value::Date(d) => d.and_hms_opt(0, 0, 0).unwrap_or_else(|| {
            chrono::NaiveDateTime::new(*d, chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap())
        }),
        Value::DateTime(dt) => *dt,
        Value::Time(t) => {
            // For Time values, use a dummy date
            let dummy_date = chrono::NaiveDate::from_ymd_opt(1899, 12, 30).unwrap();
            chrono::NaiveDateTime::new(dummy_date, *t)
        }
        _ => return value_to_string(val),
    };
    
    // Check for AM/PM pattern - affects hour format
    let has_ampm = fmt.to_lowercase().contains("am/pm") || fmt.contains("AM/PM");
    
    let mut pattern = fmt.to_string();
    
    // Handle AM/PM first (before other replacements)
    pattern = pattern.replace("AM/PM", "%p");
    pattern = pattern.replace("am/pm", "%p");
    
    // Order matters - replace longer patterns first!
    pattern = pattern.replace("yyyy", "%Y");
    pattern = pattern.replace("yy", "%y");
    pattern = pattern.replace("mmmm", "%B");  // Full month name (January)
    pattern = pattern.replace("mmm", "%b");   // Short month name (Jan)
    pattern = pattern.replace("dddd", "%A");  // Full weekday name (Monday)
    pattern = pattern.replace("ddd", "%a");   // Short weekday name (Mon)
    pattern = pattern.replace("dd", "%d");    // Day of month (2 digit)
    pattern = pattern.replace("MM", "%M");    // VBA MM in time = minutes
    pattern = pattern.replace("mm", "%m");    // VBA mm = month
    pattern = pattern.replace("HH", "%H");    // 24-hour (2 digit)
    pattern = pattern.replace("hh", if has_ampm { "%I" } else { "%H" });  // 12-hour if AM/PM present
    pattern = pattern.replace("nn", "%M");    // VBA nn = minutes (2 digit)
    pattern = pattern.replace("SS", "%S");    // Seconds (2 digit)
    pattern = pattern.replace("ss", "%S");    // Seconds (2 digit)
    
    // Handle single letter patterns (must come after multi-letter patterns)
    let chars: Vec<char> = pattern.chars().collect();
    let mut result = String::new();
    let mut i = 0;
    while i < chars.len() {
        if chars[i] == '%' && i + 1 < chars.len() {
            // Skip chrono format specifier
            result.push(chars[i]);
            result.push(chars[i + 1]);
            i += 2;
        } else if chars[i] == 'd' {
            // Single 'd' for day without leading zero
            result.push_str("%-d");
            i += 1;
        } else if chars[i] == 'm' {
            // Single 'm' for month without leading zero
            result.push_str("%-m");
            i += 1;
        } else if chars[i] == 'h' {
            // Single 'h' for hour without leading zero
            if has_ampm {
                result.push_str("%-I");  // 12-hour without leading zero
            } else {
                result.push_str("%-H");  // 24-hour without leading zero
            }
            i += 1;
        } else if chars[i] == 'n' {
            // Single 'n' for minute without leading zero
            result.push_str("%-M");
            i += 1;
        } else if chars[i] == 's' {
            // Single 's' for second without leading zero
            result.push_str("%-S");
            i += 1;
        } else {
            result.push(chars[i]);
            i += 1;
        }
    }
    pattern = result;
    
    dt.format(&pattern).to_string()
}
