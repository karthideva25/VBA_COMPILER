// src/host/excel/properties/range_properties.rs
// ============================================================================
// Property handlers for Range object
// 
// This file serves as a TEMPLATE for implementing Excel object properties.
// Each property follows this pattern:
//   1. Match the property name (case-insensitive)
//   2. Return a placeholder value or call the engine
//   3. Include a TODO comment where engine integration is needed
//
// Properties can be:
//   - Read-only: Only implement in get_range_property()
//   - Read-write: Implement in both get_range_property() and set_range_property()
//   - Write-only: Only implement in set_range_property() (rare)
//
// Note: For sub-objects (Font, Interior, Borders, etc.), we return a String
// with format "ObjectType:data" which the interpreter can parse and dispatch.
// When proper COM support is added, these should return Value::Object.
// ============================================================================

use anyhow::{Result, bail};
use crate::context::Value;
use crate::host::excel::engine;

// ============================================================================
// GET PROPERTIES
// ============================================================================

/// Get Range property by name
/// 
/// # Arguments
/// * `address` - Cell address like "A1" or range like "A1:B5"
/// * `property` - Property name (case-insensitive)
/// 
/// # Returns
/// * `Ok(Value)` - The property value
/// * `Err` - If property is unknown or engine call fails
pub fn get_range_property(address: &str, property: &str) -> Result<Value> {
    match property.to_lowercase().as_str() {
        
        // ====================================================================
        // CONTENT & VALUES
        // ====================================================================
        
        "value" => {
            // Returns the value of the cell(s)
            // TODO: For multi-cell ranges, return 2D array
            match engine::get_cell_value(address) {
                Ok(val) => {
                    if val.is_empty() {
                        Ok(Value::Empty)
                    } else if let Ok(n) = val.parse::<i64>() {
                        Ok(Value::Integer(n))
                    } else if let Ok(n) = val.parse::<f64>() {
                        Ok(Value::Double(n))
                    } else {
                        Ok(Value::String(val))
                    }
                }
                Err(e) => bail!("Failed to get cell value: {}", e),
            }
        }
        
        "value2" => {
            // Same as Value but dates are returned as serial numbers
            // TODO: ENGINE CALL - engine::get_cell_value_raw(address)
            match engine::get_cell_value(address) {
                Ok(val) => {
                    if val.is_empty() {
                        Ok(Value::Empty)
                    } else if let Ok(n) = val.parse::<i64>() {
                        Ok(Value::Integer(n))
                    } else if let Ok(n) = val.parse::<f64>() {
                        Ok(Value::Double(n))
                    } else {
                        Ok(Value::String(val))
                    }
                }
                Err(e) => bail!("Failed to get cell value: {}", e),
            }
        }
        
        "text" => {
            // Returns the formatted text as displayed in the cell
            // TODO: ENGINE CALL - engine::get_cell_formatted_text(address)
            eprintln!("   [STUB] Range({}).Text - returning raw value", address);
            match engine::get_cell_value(address) {
                Ok(val) => Ok(Value::String(val)),
                Err(e) => bail!("Failed to get cell text: {}", e),
            }
        }
        
        "formula" => {
            // Returns the formula in A1 notation (e.g., "=A1+B1")
            // TODO: ENGINE CALL - engine::get_cell_formula(address)
            eprintln!("   [STUB] Range({}).Formula - returning empty", address);
            Ok(Value::String(String::new()))
        }
        
        "formular1c1" => {
            // Returns the formula in R1C1 notation (e.g., "=R[-1]C+R[-1]C[1]")
            // TODO: ENGINE CALL - engine::get_cell_formula_r1c1(address)
            eprintln!("   [STUB] Range({}).FormulaR1C1 - returning empty", address);
            Ok(Value::String(String::new()))
        }
        
        "formulaarray" => {
            // Returns the array formula (if any)
            // TODO: ENGINE CALL - engine::get_array_formula(address)
            eprintln!("   [STUB] Range({}).FormulaArray - returning empty", address);
            Ok(Value::String(String::new()))
        }
        
        "hasarray" => {
            // Returns True if cell is part of an array formula
            // TODO: ENGINE CALL - engine::has_array_formula(address)
            eprintln!("   [STUB] Range({}).HasArray - returning False", address);
            Ok(Value::Boolean(false))
        }
        
        // ====================================================================
        // ADDRESS & LOCATION
        // ====================================================================
        
        "address" => {
            // Returns the absolute address as string (e.g., "$A$1")
            // Format: $A$1 (absolute by default)
            Ok(Value::String(format!("${}", address.to_uppercase())))
        }
        
        "row" => {
            // Returns the row number (1-based)
            let (row, _) = engine::address_to_indices(address)
                .map_err(|e| anyhow::anyhow!("{}", e))?;
            Ok(Value::Integer((row + 1) as i64))
        }
        
        "column" => {
            // Returns the column number (1-based)
            let (_, col) = engine::address_to_indices(address)
                .map_err(|e| anyhow::anyhow!("{}", e))?;
            Ok(Value::Integer((col + 1) as i64))
        }
        
        "rows" => {
            // Returns a Range representing all rows in the range
            // In VBA, Range.Rows.Count returns the number of rows
            // For now we return the range itself (Rows collection)
            // The Count property will be handled when accessed on this
            Ok(Value::String(format!("Range:{}", address)))
        }
        
        "columns" => {
            // Returns a Range representing all columns in the range
            // In VBA, Range.Columns.Count returns the number of columns
            Ok(Value::String(format!("Range:{}", address)))
        }
        
        "cells" => {
            // Returns a Range representing all cells in the range
            // In VBA, Range.Cells can be indexed like Range.Cells(1,1)
            // For direct property access, return self
            Ok(Value::String(format!("Range:{}", address)))
        }
        
        "entirerow" => {
            // Returns entire row(s) containing the range
            // TODO: ENGINE CALL - engine::get_entire_row(address)
            let (row, _) = engine::address_to_indices(address)
                .map_err(|e| anyhow::anyhow!("{}", e))?;
            let entire_row = format!("{}:{}", row + 1, row + 1);
            eprintln!("   [STUB] Range({}).EntireRow -> {}", address, entire_row);
            Ok(Value::String(format!("Range:{}", entire_row)))
        }
        
        "entirecolumn" => {
            // Returns entire column(s) containing the range
            // TODO: ENGINE CALL - engine::get_entire_column(address)
            let (_, col) = engine::address_to_indices(address)
                .map_err(|e| anyhow::anyhow!("{}", e))?;
            let col_letter = column_index_to_letter(col);
            let entire_col = format!("{}:{}", col_letter, col_letter);
            eprintln!("   [STUB] Range({}).EntireColumn -> {}", address, entire_col);
            Ok(Value::String(format!("Range:{}", entire_col)))
        }
        
        "currentregion" => {
            // Returns the current region (bounded by empty rows/columns)
            // TODO: ENGINE CALL - engine::get_current_region(address)
            eprintln!("   [STUB] Range({}).CurrentRegion - returning self", address);
            Ok(Value::String(format!("Range:{}", address)))
        }
        
        "areas" => {
            // Returns an Areas collection for non-contiguous ranges
            // TODO: ENGINE CALL - engine::get_range_areas(address)
            eprintln!("   [STUB] Range({}).Areas - returning self", address);
            Ok(Value::String(format!("Areas:{}", address)))
        }
        
        // ====================================================================
        // COUNT & SIZE
        // ====================================================================
        
        "count" => {
            // Returns the number of cells in the range (as Long)
            let (rows, cols) = get_range_dimensions(address)?;
            let count = rows as i64 * cols as i64;
            Ok(Value::Integer(count))
        }
        
        "countlarge" => {
            // Returns the number of cells (as Double, for large ranges)
            let (rows, cols) = get_range_dimensions(address)?;
            let count = rows as f64 * cols as f64;
            Ok(Value::Double(count))
        }
        
        // ====================================================================
        // FORMATTING - NUMBER
        // ====================================================================
        
        "numberformat" => {
            // Returns the number format code (e.g., "0.00", "@", "General")
            // TODO: ENGINE CALL - engine::get_cell_number_format(address)
            eprintln!("   [STUB] Range({}).NumberFormat - returning 'General'", address);
            Ok(Value::String("General".to_string()))
        }
        
        // ====================================================================
        // FORMATTING - FONT (Sub-object)
        // ====================================================================
        
        "font" => {
            // Returns a Font object for font formatting
            // The interpreter should handle Font.Name, Font.Bold, etc.
            // TODO: Return proper Font object reference when COM support is added
            eprintln!("   [STUB] Range({}).Font - returning Font object reference", address);
            Ok(Value::String(format!("Font:{}", address)))
        }
        
        // ====================================================================
        // FORMATTING - INTERIOR (Sub-object)
        // ====================================================================
        
        "interior" => {
            // Returns an Interior object for fill/background
            // The interpreter should handle Interior.Color, Interior.Pattern, etc.
            // TODO: Return proper Interior object reference when COM support is added
            eprintln!("   [STUB] Range({}).Interior - returning Interior object reference", address);
            Ok(Value::String(format!("Interior:{}", address)))
        }
        
        // ====================================================================
        // FORMATTING - BORDERS (Sub-object)
        // ====================================================================
        
        "borders" => {
            // Returns a Borders collection for cell borders
            // The interpreter should handle Borders(xlEdgeLeft), etc.
            // TODO: Return proper Borders object reference when COM support is added
            eprintln!("   [STUB] Range({}).Borders - returning Borders object reference", address);
            Ok(Value::String(format!("Borders:{}", address)))
        }
        
        // ====================================================================
        // FORMATTING - ALIGNMENT
        // ====================================================================
        
        "horizontalalignment" => {
            // Returns horizontal alignment (xlLeft, xlCenter, xlRight, etc.)
            // TODO: ENGINE CALL - engine::get_horizontal_alignment(address)
            eprintln!("   [STUB] Range({}).HorizontalAlignment - returning xlGeneral (-4105)", address);
            Ok(Value::Integer(-4105)) // xlGeneral
        }
        
        "verticalalignment" => {
            // Returns vertical alignment (xlTop, xlCenter, xlBottom, etc.)
            // TODO: ENGINE CALL - engine::get_vertical_alignment(address)
            eprintln!("   [STUB] Range({}).VerticalAlignment - returning xlBottom (-4107)", address);
            Ok(Value::Integer(-4107)) // xlBottom
        }
        
        "orientation" => {
            // Returns text orientation in degrees (-90 to 90) or xlVertical
            // TODO: ENGINE CALL - engine::get_text_orientation(address)
            eprintln!("   [STUB] Range({}).Orientation - returning 0", address);
            Ok(Value::Integer(0))
        }
        
        "wraptext" => {
            // Returns True if text wrapping is enabled
            // TODO: ENGINE CALL - engine::get_wrap_text(address)
            eprintln!("   [STUB] Range({}).WrapText - returning False", address);
            Ok(Value::Boolean(false))
        }
        
        "addindent" => {
            // Returns True if text is indented when alignment is set
            // TODO: ENGINE CALL - engine::get_add_indent(address)
            eprintln!("   [STUB] Range({}).AddIndent - returning False", address);
            Ok(Value::Boolean(false))
        }
        
        "indentlevel" => {
            // Returns the indent level (0-15)
            // TODO: ENGINE CALL - engine::get_indent_level(address)
            eprintln!("   [STUB] Range({}).IndentLevel - returning 0", address);
            Ok(Value::Integer(0))
        }
        
        // ====================================================================
        // CELL STATE & PROTECTION
        // ====================================================================
        
        "locked" => {
            // Returns True if cells are locked
            // TODO: ENGINE CALL - engine::get_cell_locked(address)
            eprintln!("   [STUB] Range({}).Locked - returning True", address);
            Ok(Value::Boolean(true)) // Default is locked
        }
        
        "hidden" => {
            // Returns True if rows/columns containing range are hidden
            // TODO: ENGINE CALL - engine::get_range_hidden(address)
            eprintln!("   [STUB] Range({}).Hidden - returning False", address);
            Ok(Value::Boolean(false))
        }
        
        "mergecells" => {
            // Returns True if range is part of a merged cell
            // TODO: ENGINE CALL - engine::get_merge_cells(address)
            eprintln!("   [STUB] Range({}).MergeCells - returning False", address);
            Ok(Value::Boolean(false))
        }
        
        // ====================================================================
        // DEPENDENCIES & PRECEDENTS
        // ====================================================================
        
        "dependents" => {
            // Returns a Range of all dependent cells (direct and indirect)
            // TODO: ENGINE CALL - engine::get_dependents(address)
            eprintln!("   [STUB] Range({}).Dependents - returning Nothing", address);
            Ok(Value::Empty)
        }
        
        "precedents" => {
            // Returns a Range of all precedent cells (direct and indirect)
            // TODO: ENGINE CALL - engine::get_precedents(address)
            eprintln!("   [STUB] Range({}).Precedents - returning Nothing", address);
            Ok(Value::Empty)
        }
        
        "directdependents" => {
            // Returns a Range of directly dependent cells only
            // TODO: ENGINE CALL - engine::get_direct_dependents(address)
            eprintln!("   [STUB] Range({}).DirectDependents - returning Nothing", address);
            Ok(Value::Empty)
        }
        
        "directprecedents" => {
            // Returns a Range of directly precedent cells only
            // TODO: ENGINE CALL - engine::get_direct_precedents(address)
            eprintln!("   [STUB] Range({}).DirectPrecedents - returning Nothing", address);
            Ok(Value::Empty)
        }
        
        "specialcells" => {
            // Returns a Range of cells matching special criteria
            // Note: This is typically called as a method with arguments
            // TODO: ENGINE CALL - engine::get_special_cells(address, type, value)
            eprintln!("   [STUB] Range({}).SpecialCells - returning self", address);
            Ok(Value::String(format!("Range:{}", address)))
        }
        
        // ====================================================================
        // STYLE & NAMED ITEMS
        // ====================================================================
        
        "style" => {
            // Returns the Style object applied to the range
            // TODO: ENGINE CALL - engine::get_cell_style(address)
            eprintln!("   [STUB] Range({}).Style - returning 'Normal'", address);
            Ok(Value::String("Normal".to_string()))
        }
        
        "name" => {
            // Returns the Name object if range has a defined name
            // TODO: ENGINE CALL - engine::get_range_name(address)
            eprintln!("   [STUB] Range({}).Name - returning Nothing", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // COMMENTS & ANNOTATIONS
        // ====================================================================
        
        "comment" => {
            // Returns the Comment object (if any)
            // TODO: ENGINE CALL - engine::get_cell_comment(address)
            eprintln!("   [STUB] Range({}).Comment - returning Nothing", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // HYPERLINKS & VALIDATION
        // ====================================================================
        
        "hyperlinks" => {
            // Returns the Hyperlinks collection for the range
            // TODO: ENGINE CALL - engine::get_hyperlinks(address)
            eprintln!("   [STUB] Range({}).Hyperlinks - returning Hyperlinks object reference", address);
            Ok(Value::String(format!("Hyperlinks:{}", address)))
        }
        
        "validation" => {
            // Returns the Validation object (data validation settings)
            // TODO: ENGINE CALL - engine::get_validation(address)
            eprintln!("   [STUB] Range({}).Validation - returning Validation object reference", address);
            Ok(Value::String(format!("Validation:{}", address)))
        }
        
        // ====================================================================
        // OBJECT MODEL
        // ====================================================================
        
        "creator" => {
            // Returns the creator application (always Excel = 1480803660)
            Ok(Value::Integer(1480803660)) // xlCreatorCode for Excel
        }
        
        "parent" => {
            // Returns the parent Worksheet object
            // TODO: Return proper Worksheet reference when COM support is added
            eprintln!("   [STUB] Range({}).Parent - returning Worksheet object reference", address);
            Ok(Value::String(format!("Worksheet:{}", engine::get_active_sheet())))
        }
        
        // ====================================================================
        // UNKNOWN PROPERTY
        // ====================================================================
        
        _ => bail!("Unknown Range property: {}", property),
    }
}

// ============================================================================
// SET PROPERTIES
// ============================================================================

/// Set Range property by name
/// 
/// # Arguments
/// * `address` - Cell address like "A1" or range like "A1:B5"
/// * `property` - Property name (case-insensitive)
/// * `value` - Value to set
/// 
/// # Returns
/// * `Ok(())` - Property was set successfully
/// * `Err` - If property is read-only, unknown, or engine call fails
pub fn set_range_property(address: &str, property: &str, value: Value) -> Result<()> {
    match property.to_lowercase().as_str() {
        
        // ====================================================================
        // CONTENT & VALUES
        // ====================================================================
        
        "value" | "value2" => {
            // Set the value of the cell(s)
            let value_str = value_to_string(&value);
            engine::set_cell_value(address, &value_str)
                .map_err(|e| anyhow::anyhow!("Failed to set cell value: {}", e))
        }
        
        "formula" => {
            // Set formula in A1 notation
            // TODO: ENGINE CALL - engine::set_cell_formula(address, formula)
            let formula = value_to_string(&value);
            eprintln!("   [STUB] Range({}).Formula = '{}' - storing as value", address, formula);
            engine::set_cell_value(address, &formula)
                .map_err(|e| anyhow::anyhow!("Failed to set formula: {}", e))
        }
        
        "formular1c1" => {
            // Set formula in R1C1 notation
            // TODO: ENGINE CALL - engine::set_cell_formula_r1c1(address, formula)
            let formula = value_to_string(&value);
            eprintln!("   [STUB] Range({}).FormulaR1C1 = '{}' - NOT IMPLEMENTED", address, formula);
            Ok(())
        }
        
        "formulaarray" => {
            // Set array formula
            // TODO: ENGINE CALL - engine::set_array_formula(address, formula)
            let formula = value_to_string(&value);
            eprintln!("   [STUB] Range({}).FormulaArray = '{}' - NOT IMPLEMENTED", address, formula);
            Ok(())
        }
        
        // ====================================================================
        // FORMATTING - NUMBER
        // ====================================================================
        
        "numberformat" => {
            // Set number format code
            // TODO: ENGINE CALL - engine::set_cell_number_format(address, format)
            let format = value_to_string(&value);
            eprintln!("   [STUB] Range({}).NumberFormat = '{}' - NOT IMPLEMENTED", address, format);
            Ok(())
        }
        
        // ====================================================================
        // FORMATTING - ALIGNMENT
        // ====================================================================
        
        "horizontalalignment" => {
            // Set horizontal alignment
            // TODO: ENGINE CALL - engine::set_horizontal_alignment(address, align)
            eprintln!("   [STUB] Range({}).HorizontalAlignment = {:?} - NOT IMPLEMENTED", address, value);
            Ok(())
        }
        
        "verticalalignment" => {
            // Set vertical alignment
            // TODO: ENGINE CALL - engine::set_vertical_alignment(address, align)
            eprintln!("   [STUB] Range({}).VerticalAlignment = {:?} - NOT IMPLEMENTED", address, value);
            Ok(())
        }
        
        "orientation" => {
            // Set text orientation
            // TODO: ENGINE CALL - engine::set_text_orientation(address, degrees)
            eprintln!("   [STUB] Range({}).Orientation = {:?} - NOT IMPLEMENTED", address, value);
            Ok(())
        }
        
        "wraptext" => {
            // Set text wrapping
            // TODO: ENGINE CALL - engine::set_wrap_text(address, wrap)
            let wrap = value_to_bool(&value);
            eprintln!("   [STUB] Range({}).WrapText = {} - NOT IMPLEMENTED", address, wrap);
            Ok(())
        }
        
        "addindent" => {
            // Set add indent flag
            // TODO: ENGINE CALL - engine::set_add_indent(address, add_indent)
            let add_indent = value_to_bool(&value);
            eprintln!("   [STUB] Range({}).AddIndent = {} - NOT IMPLEMENTED", address, add_indent);
            Ok(())
        }
        
        "indentlevel" => {
            // Set indent level (0-15)
            // TODO: ENGINE CALL - engine::set_indent_level(address, level)
            let level = value_to_int(&value);
            eprintln!("   [STUB] Range({}).IndentLevel = {} - NOT IMPLEMENTED", address, level);
            Ok(())
        }
        
        // ====================================================================
        // CELL STATE & PROTECTION
        // ====================================================================
        
        "locked" => {
            // Set locked state
            // TODO: ENGINE CALL - engine::set_cell_locked(address, locked)
            let locked = value_to_bool(&value);
            eprintln!("   [STUB] Range({}).Locked = {} - NOT IMPLEMENTED", address, locked);
            Ok(())
        }
        
        "hidden" => {
            // Set hidden state (for rows/columns)
            // TODO: ENGINE CALL - engine::set_range_hidden(address, hidden)
            let hidden = value_to_bool(&value);
            eprintln!("   [STUB] Range({}).Hidden = {} - NOT IMPLEMENTED", address, hidden);
            Ok(())
        }
        
        "mergecells" => {
            // Set merge state (True to merge, False to unmerge)
            // TODO: ENGINE CALL - engine::set_merge_cells(address, merge)
            let merge = value_to_bool(&value);
            eprintln!("   [STUB] Range({}).MergeCells = {} - NOT IMPLEMENTED", address, merge);
            Ok(())
        }
        
        // ====================================================================
        // STYLE
        // ====================================================================
        
        "style" => {
            // Apply a named style
            // TODO: ENGINE CALL - engine::set_cell_style(address, style_name)
            let style = value_to_string(&value);
            eprintln!("   [STUB] Range({}).Style = '{}' - NOT IMPLEMENTED", address, style);
            Ok(())
        }
        
        "name" => {
            // Create a defined name for this range
            // TODO: ENGINE CALL - engine::create_range_name(address, name)
            let name = value_to_string(&value);
            eprintln!("   [STUB] Range({}).Name = '{}' - NOT IMPLEMENTED", address, name);
            Ok(())
        }
        
        // ====================================================================
        // READ-ONLY PROPERTIES (return error)
        // ====================================================================
        
        "text" | "address" | "row" | "column" | "rows" | "columns" | "cells" |
        "entirerow" | "entirecolumn" | "currentregion" | "areas" |
        "count" | "countlarge" | "hasarray" |
        "font" | "interior" | "borders" |
        "dependents" | "precedents" | "directdependents" | "directprecedents" |
        "specialcells" | "comment" | "hyperlinks" | "validation" |
        "creator" | "parent" => {
            bail!("Range.{} is read-only", property)
        }
        
        // ====================================================================
        // UNKNOWN PROPERTY
        // ====================================================================
        
        _ => bail!("Cannot set Range property: {}", property),
    }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/// Parse a range address and return (row_count, col_count)
/// Handles both single cell (e.g., "A1") and range (e.g., "A1:C5")
fn get_range_dimensions(address: &str) -> Result<(i32, i32)> {
    if let Some(colon_pos) = address.find(':') {
        // Range like "A1:C5"
        let start = &address[..colon_pos];
        let end = &address[colon_pos + 1..];
        
        let (start_row, start_col) = engine::address_to_indices(start)
            .map_err(|e| anyhow::anyhow!("{}", e))?;
        let (end_row, end_col) = engine::address_to_indices(end)
            .map_err(|e| anyhow::anyhow!("{}", e))?;
        
        let rows = (end_row - start_row + 1).abs();
        let cols = (end_col - start_col + 1).abs();
        
        Ok((rows, cols))
    } else {
        // Single cell like "A1"
        Ok((1, 1))
    }
}

/// Get the start and end indices of a range
/// Returns ((start_row, start_col), (end_row, end_col))
fn get_range_bounds(address: &str) -> Result<((i32, i32), (i32, i32))> {
    if let Some(colon_pos) = address.find(':') {
        let start = &address[..colon_pos];
        let end = &address[colon_pos + 1..];
        
        let start_pos = engine::address_to_indices(start)
            .map_err(|e| anyhow::anyhow!("{}", e))?;
        let end_pos = engine::address_to_indices(end)
            .map_err(|e| anyhow::anyhow!("{}", e))?;
        
        Ok((start_pos, end_pos))
    } else {
        let pos = engine::address_to_indices(address)
            .map_err(|e| anyhow::anyhow!("{}", e))?;
        Ok((pos, pos))
    }
}

/// Convert (row, col) to Excel address
fn indices_to_address(row: i32, col: i32) -> String {
    format!("{}{}", column_index_to_letter(col), row + 1)
}

/// Convert column index (0-based) to Excel column letter (A, B, ..., Z, AA, AB, ...)
fn column_index_to_letter(col: i32) -> String {
    let mut result = String::new();
    let mut col = col + 1; // Convert to 1-based
    while col > 0 {
        let remainder = ((col - 1) % 26) as u8;
        result.insert(0, (b'A' + remainder) as char);
        col = (col - 1) / 26;
    }
    result
}

/// Convert Value to String representation
fn value_to_string(value: &Value) -> String {
    match value {
        Value::String(s) => s.clone(),
        Value::Integer(i) => i.to_string(),
        Value::Double(d) => d.to_string(),
        Value::Boolean(b) => if *b { "TRUE".to_string() } else { "FALSE".to_string() },
        Value::Currency(c) => c.to_string(),
        Value::Empty => String::new(),
        other => other.as_string(),
    }
}

/// Convert Value to bool
fn value_to_bool(value: &Value) -> bool {
    match value {
        Value::Boolean(b) => *b,
        Value::Integer(i) => *i != 0,
        Value::Double(d) => *d != 0.0,
        Value::String(s) => s.eq_ignore_ascii_case("true") || s == "1",
        _ => false,
    }
}

/// Convert Value to i64
fn value_to_int(value: &Value) -> i64 {
    match value {
        Value::Integer(i) => *i,
        Value::Double(d) => *d as i64,
        Value::Boolean(b) => if *b { 1 } else { 0 },
        Value::String(s) => s.parse().unwrap_or(0),
        _ => 0,
    }
}

// ============================================================================
// TESTS
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_column_index_to_letter() {
        assert_eq!(column_index_to_letter(0), "A");
        assert_eq!(column_index_to_letter(1), "B");
        assert_eq!(column_index_to_letter(25), "Z");
        assert_eq!(column_index_to_letter(26), "AA");
        assert_eq!(column_index_to_letter(27), "AB");
        assert_eq!(column_index_to_letter(701), "ZZ");
        assert_eq!(column_index_to_letter(702), "AAA");
    }
    
    #[test]
    fn test_value_to_string() {
        assert_eq!(value_to_string(&Value::String("test".to_string())), "test");
        assert_eq!(value_to_string(&Value::Integer(42)), "42");
        assert_eq!(value_to_string(&Value::Boolean(true)), "TRUE");
        assert_eq!(value_to_string(&Value::Empty), "");
    }
}
