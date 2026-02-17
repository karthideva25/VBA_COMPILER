// src/host/excel/methods/range_methods.rs
// ============================================================================
// Method handlers for Range object
// 
// This file serves as a TEMPLATE for implementing Excel object methods.
// Each method follows this pattern:
//   1. Match the method name (case-insensitive)
//   2. Print a stub message or perform the action
//   3. Return appropriate Value (usually Value::Empty for void methods)
//   4. Include a TODO comment where engine integration is needed
//
// Methods can:
//   - Return nothing (Sub-like): Return Value::Empty
//   - Return a value (Function-like): Return the actual Value
//   - Return an object (Factory): Return Value::String("ObjectType:data")
//
// Arguments are passed as &[Value] - check args.len() and extract with pattern matching.
// ============================================================================

use anyhow::{Result, bail};
use crate::context::Value;
use crate::host::excel::engine;

// ============================================================================
// CALL METHOD
// ============================================================================

/// Call method on Range object
/// 
/// # Arguments
/// * `address` - Cell address like "A1" or range like "A1:B5"
/// * `method` - Method name (case-insensitive)
/// * `args` - Method arguments (may be empty)
/// 
/// # Returns
/// * `Ok(Value)` - The method return value (often Value::Empty for void methods)
/// * `Err` - If method is unknown or engine call fails
pub fn call_range_method(address: &str, method: &str, args: &[Value]) -> Result<Value> {
    match method.to_lowercase().as_str() {
        
        // ====================================================================
        // SELECTION & ACTIVATION
        // ====================================================================
        
        "select" => {
            // Selects the range (makes it the current selection)
            // TODO: ENGINE CALL - engine::select_range(address)
            eprintln!("   [STUB] Range({}).Select()", address);
            Ok(Value::Empty)
        }
        
        "activate" => {
            // Activates a single cell within a selection
            // TODO: ENGINE CALL - engine::activate_cell(address)
            eprintln!("   [STUB] Range({}).Activate()", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // CLIPBOARD OPERATIONS
        // ====================================================================
        
        "copy" => {
            // Copy([Destination])
            // Copies the range to clipboard or to Destination if specified
            // TODO: ENGINE CALL - engine::copy_range(address, destination)
            if let Some(dest) = args.first() {
                let dest_addr = value_to_string(dest);
                eprintln!("   [STUB] Range({}).Copy(Destination:={})", address, dest_addr);
            } else {
                eprintln!("   [STUB] Range({}).Copy() - to clipboard", address);
            }
            Ok(Value::Empty)
        }
        
        "cut" => {
            // Cut([Destination])
            // Cuts the range to clipboard or moves to Destination if specified
            // TODO: ENGINE CALL - engine::cut_range(address, destination)
            if let Some(dest) = args.first() {
                let dest_addr = value_to_string(dest);
                eprintln!("   [STUB] Range({}).Cut(Destination:={})", address, dest_addr);
            } else {
                eprintln!("   [STUB] Range({}).Cut() - to clipboard", address);
            }
            Ok(Value::Empty)
        }
        
        "pastespecial" => {
            // PasteSpecial([Paste], [Operation], [SkipBlanks], [Transpose])
            // Pastes from clipboard with special options
            // Paste: xlPasteAll(-4104), xlPasteValues(-4163), xlPasteFormulas(-4123), etc.
            // TODO: ENGINE CALL - engine::paste_special(address, paste_type, operation, skip_blanks, transpose)
            let paste_type = args.get(0).map(value_to_int).unwrap_or(-4104); // xlPasteAll
            let operation = args.get(1).map(value_to_int).unwrap_or(-4142);  // xlNone
            let skip_blanks = args.get(2).map(value_to_bool).unwrap_or(false);
            let transpose = args.get(3).map(value_to_bool).unwrap_or(false);
            eprintln!("   [STUB] Range({}).PasteSpecial(Paste:={}, Operation:={}, SkipBlanks:={}, Transpose:={})", 
                     address, paste_type, operation, skip_blanks, transpose);
            Ok(Value::Empty)
        }
        
        "copypicture" => {
            // CopyPicture([Appearance], [Format])
            // Copies the range as a picture to clipboard
            // Appearance: xlScreen(1), xlPrinter(2)
            // Format: xlPicture(-4147), xlBitmap(2)
            // TODO: ENGINE CALL - engine::copy_picture(address, appearance, format)
            let appearance = args.get(0).map(value_to_int).unwrap_or(1); // xlScreen
            let format = args.get(1).map(value_to_int).unwrap_or(-4147); // xlPicture
            eprintln!("   [STUB] Range({}).CopyPicture(Appearance:={}, Format:={})", address, appearance, format);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // CLEAR OPERATIONS
        // ====================================================================
        
        "clear" => {
            // Clears everything (values, formats, comments, etc.)
            // TODO: ENGINE CALL - engine::clear_range(address)
            eprintln!("   [STUB] Range({}).Clear()", address);
            engine::set_cell_value(address, "")
                .map_err(|e| anyhow::anyhow!("Failed to clear: {}", e))?;
            Ok(Value::Empty)
        }
        
        "clearcontents" => {
            // Clears only values and formulas (keeps formatting)
            // TODO: ENGINE CALL - engine::clear_contents(address)
            eprintln!("   [STUB] Range({}).ClearContents()", address);
            engine::set_cell_value(address, "")
                .map_err(|e| anyhow::anyhow!("Failed to clear contents: {}", e))?;
            Ok(Value::Empty)
        }
        
        "clearformats" => {
            // Clears only formatting (keeps values)
            // TODO: ENGINE CALL - engine::clear_formats(address)
            eprintln!("   [STUB] Range({}).ClearFormats()", address);
            Ok(Value::Empty)
        }
        
        "clearcomments" => {
            // Clears only comments
            // TODO: ENGINE CALL - engine::clear_comments(address)
            eprintln!("   [STUB] Range({}).ClearComments()", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // FIND & REPLACE
        // ====================================================================
        
        "find" => {
            // Find(What, [After], [LookIn], [LookAt], [SearchOrder], [SearchDirection], [MatchCase], [MatchByte], [SearchFormat])
            // Returns Range of first match or Nothing
            // TODO: ENGINE CALL - engine::find_in_range(address, what, options...)
            let what = args.get(0).map(value_to_string).unwrap_or_default();
            eprintln!("   [STUB] Range({}).Find(What:='{}')", address, what);
            // Return Nothing for now (not found)
            Ok(Value::Empty)
        }
        
        "findnext" => {
            // FindNext([After])
            // Continues a Find operation
            // TODO: ENGINE CALL - engine::find_next(address, after)
            eprintln!("   [STUB] Range({}).FindNext()", address);
            Ok(Value::Empty)
        }
        
        "findprevious" => {
            // FindPrevious([After])
            // Continues a Find operation in reverse
            // TODO: ENGINE CALL - engine::find_previous(address, after)
            eprintln!("   [STUB] Range({}).FindPrevious()", address);
            Ok(Value::Empty)
        }
        
        "replace" => {
            // Replace(What, Replacement, [LookAt], [SearchOrder], [MatchCase], [MatchByte], [SearchFormat], [ReplaceFormat])
            // Returns True if replacements were made
            // TODO: ENGINE CALL - engine::replace_in_range(address, what, replacement, options...)
            let what = args.get(0).map(value_to_string).unwrap_or_default();
            let replacement = args.get(1).map(value_to_string).unwrap_or_default();
            eprintln!("   [STUB] Range({}).Replace(What:='{}', Replacement:='{}')", address, what, replacement);
            Ok(Value::Boolean(false)) // No replacements made
        }
        
        // ====================================================================
        // INSERT & DELETE
        // ====================================================================
        
        "insert" => {
            // Insert([Shift], [CopyOrigin])
            // Inserts cells, shifting existing cells
            // Shift: xlShiftDown(-4121), xlShiftToRight(-4161)
            // CopyOrigin: xlFormatFromLeftOrAbove(0), xlFormatFromRightOrBelow(1)
            // TODO: ENGINE CALL - engine::insert_cells(address, shift, copy_origin)
            let shift = args.get(0).map(value_to_int).unwrap_or(-4121); // xlShiftDown
            eprintln!("   [STUB] Range({}).Insert(Shift:={})", address, shift);
            Ok(Value::Empty)
        }
        
        "delete" => {
            // Delete([Shift])
            // Deletes cells, shifting remaining cells
            // Shift: xlShiftUp(-4162), xlShiftToLeft(-4159)
            // TODO: ENGINE CALL - engine::delete_cells(address, shift)
            let shift = args.get(0).map(value_to_int).unwrap_or(-4162); // xlShiftUp
            eprintln!("   [STUB] Range({}).Delete(Shift:={})", address, shift);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // FILL OPERATIONS
        // ====================================================================
        
        "autofill" => {
            // AutoFill(Destination, [Type])
            // Auto-fills the destination range based on this range
            // Type: xlFillDefault(0), xlFillCopy(1), xlFillSeries(2), etc.
            // TODO: ENGINE CALL - engine::auto_fill(address, destination, fill_type)
            let destination = args.get(0).map(value_to_string).unwrap_or_default();
            let fill_type = args.get(1).map(value_to_int).unwrap_or(0); // xlFillDefault
            eprintln!("   [STUB] Range({}).AutoFill(Destination:={}, Type:={})", address, destination, fill_type);
            Ok(Value::Empty)
        }
        
        "filldown" => {
            // Fills down from top cell(s) to bottom of range
            // TODO: ENGINE CALL - engine::fill_down(address)
            eprintln!("   [STUB] Range({}).FillDown()", address);
            Ok(Value::Empty)
        }
        
        "fillup" => {
            // Fills up from bottom cell(s) to top of range
            // TODO: ENGINE CALL - engine::fill_up(address)
            eprintln!("   [STUB] Range({}).FillUp()", address);
            Ok(Value::Empty)
        }
        
        "fillleft" => {
            // Fills left from right cell(s) to left of range
            // TODO: ENGINE CALL - engine::fill_left(address)
            eprintln!("   [STUB] Range({}).FillLeft()", address);
            Ok(Value::Empty)
        }
        
        "fillright" => {
            // Fills right from left cell(s) to right of range
            // TODO: ENGINE CALL - engine::fill_right(address)
            eprintln!("   [STUB] Range({}).FillRight()", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // FILTER & SORT
        // ====================================================================
        
        "autofilter" => {
            // AutoFilter([Field], [Criteria1], [Operator], [Criteria2], [VisibleDropDown])
            // Applies or removes AutoFilter
            // TODO: ENGINE CALL - engine::auto_filter(address, field, criteria1, operator, criteria2, visible_dropdown)
            let field = args.get(0).map(value_to_int);
            let criteria1 = args.get(1).map(value_to_string);
            eprintln!("   [STUB] Range({}).AutoFilter(Field:={:?}, Criteria1:={:?})", address, field, criteria1);
            Ok(Value::Empty)
        }
        
        "sort" => {
            // Sort([Key1], [Order1], [Key2], [Type], [Order2], [Key3], [Order3], [Header], [OrderCustom], [MatchCase], [Orientation], [SortMethod], [DataOption1], [DataOption2], [DataOption3])
            // Sorts the range
            // TODO: ENGINE CALL - engine::sort_range(address, key1, order1, ...)
            eprintln!("   [STUB] Range({}).Sort() - complex sort operation", address);
            Ok(Value::Empty)
        }
        
        "removeduplicates" => {
            // RemoveDuplicates([Columns], [Header])
            // Removes duplicate rows from the range
            // TODO: ENGINE CALL - engine::remove_duplicates(address, columns, header)
            eprintln!("   [STUB] Range({}).RemoveDuplicates()", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // CALCULATION
        // ====================================================================
        
        "calculate" => {
            // Forces calculation of the range
            // TODO: ENGINE CALL - engine::calculate_range(address)
            eprintln!("   [STUB] Range({}).Calculate()", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // MERGE & UNMERGE
        // ====================================================================
        
        "merge" => {
            // Merge([Across])
            // Merges cells into one merged cell
            // Across: If True, merges each row separately
            // TODO: ENGINE CALL - engine::merge_cells(address, across)
            let across = args.get(0).map(value_to_bool).unwrap_or(false);
            eprintln!("   [STUB] Range({}).Merge(Across:={})", address, across);
            Ok(Value::Empty)
        }
        
        "unmerge" => {
            // Unmerges merged cells back to individual cells
            // TODO: ENGINE CALL - engine::unmerge_cells(address)
            eprintln!("   [STUB] Range({}).UnMerge()", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // GROUPING
        // ====================================================================
        
        "group" => {
            // Groups rows or columns for outlining
            // TODO: ENGINE CALL - engine::group_range(address)
            eprintln!("   [STUB] Range({}).Group()", address);
            Ok(Value::Empty)
        }
        
        "ungroup" => {
            // Ungroups rows or columns from outlining
            // TODO: ENGINE CALL - engine::ungroup_range(address)
            eprintln!("   [STUB] Range({}).Ungroup()", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // CONSOLIDATION
        // ====================================================================
        
        "consolidate" => {
            // Consolidate([Sources], [Function], [TopRow], [LeftColumn], [CreateLinks])
            // Consolidates data from multiple ranges
            // TODO: ENGINE CALL - engine::consolidate(address, sources, function, ...)
            eprintln!("   [STUB] Range({}).Consolidate()", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // DEPENDENCIES VISUALIZATION
        // ====================================================================
        
        "showdependents" => {
            // ShowDependents([Remove])
            // Shows tracer arrows to dependents
            // Remove: If True, removes the arrows
            // TODO: ENGINE CALL - engine::show_dependents(address, remove)
            let remove = args.get(0).map(value_to_bool).unwrap_or(false);
            eprintln!("   [STUB] Range({}).ShowDependents(Remove:={})", address, remove);
            Ok(Value::Empty)
        }
        
        "showprecedents" => {
            // ShowPrecedents([Remove])
            // Shows tracer arrows to precedents
            // Remove: If True, removes the arrows
            // TODO: ENGINE CALL - engine::show_precedents(address, remove)
            let remove = args.get(0).map(value_to_bool).unwrap_or(false);
            eprintln!("   [STUB] Range({}).ShowPrecedents(Remove:={})", address, remove);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // PIVOT TABLE
        // ====================================================================
        
        "pivottablewizard" => {
            // PivotTableWizard([SourceType], [SourceData], [TableDestination], [TableName], [RowGrand], [ColumnGrand], [SaveData], [HasAutoFormat], [AutoPage], [Reserved], [BackgroundQuery], [OptimizeCache], [PageFieldOrder], [PageFieldWrapCount], [ReadData], [Connection])
            // Creates a PivotTable from this range
            // TODO: ENGINE CALL - engine::create_pivot_table(address, ...)
            eprintln!("   [STUB] Range({}).PivotTableWizard() - complex operation", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // COMMENTS
        // ====================================================================
        
        "addcomment" => {
            // AddComment([Text])
            // Adds a comment to the cell
            // Returns the Comment object
            // TODO: ENGINE CALL - engine::add_comment(address, text)
            let text = args.get(0).map(value_to_string).unwrap_or_default();
            eprintln!("   [STUB] Range({}).AddComment(Text:='{}')", address, text);
            // Return reference to Comment object
            Ok(Value::String(format!("Comment:{}", address)))
        }
        
        "clearcomment" => {
            // Clears the comment (alias for ClearComments for single cell)
            // TODO: ENGINE CALL - engine::clear_comment(address)
            eprintln!("   [STUB] Range({}).ClearComment()", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // SPECIAL CELLS
        // ====================================================================
        
        "specialcells" => {
            // SpecialCells(Type, [Value])
            // Returns cells matching special criteria
            // Type: xlCellTypeConstants(2), xlCellTypeFormulas(-4123), xlCellTypeBlanks(4), etc.
            // Value: xlNumbers(1), xlTextValues(2), xlLogical(4), xlErrors(16)
            // TODO: ENGINE CALL - engine::get_special_cells(address, type, value)
            let cell_type = args.get(0).map(value_to_int).unwrap_or(2); // xlCellTypeConstants
            let value_type = args.get(1).map(value_to_int);
            eprintln!("   [STUB] Range({}).SpecialCells(Type:={}, Value:={:?})", address, cell_type, value_type);
            // Return self for now
            Ok(Value::String(format!("Range:{}", address)))
        }
        
        // ====================================================================
        // OFFSET & RESIZE (These are often treated as properties but can be methods)
        // ====================================================================
        
        "offset" => {
            // Offset(RowOffset, ColumnOffset)
            // Returns a Range offset by the specified rows and columns
            // The returned range has the SAME size as the original
            let row_offset = args.get(0).map(value_to_int).unwrap_or(0) as i32;
            let col_offset = args.get(1).map(value_to_int).unwrap_or(0) as i32;
            
            // Parse the range to get bounds
            let ((start_row, start_col), (end_row, end_col)) = get_range_bounds(address)?;
            
            // Apply offset to both start and end
            let new_start_row = start_row + row_offset;
            let new_start_col = start_col + col_offset;
            let new_end_row = end_row + row_offset;
            let new_end_col = end_col + col_offset;
            
            // Check for negative indices
            if new_start_row < 0 || new_start_col < 0 {
                bail!("Offset would result in negative row/column indices");
            }
            
            // Create the new address
            let new_address = if start_row == end_row && start_col == end_col {
                // Single cell
                indices_to_address(new_start_row, new_start_col)
            } else {
                // Multi-cell range
                format!("{}:{}", 
                    indices_to_address(new_start_row, new_start_col),
                    indices_to_address(new_end_row, new_end_col))
            };
            
            Ok(Value::String(format!("Range:{}", new_address)))
        }
        
        "resize" => {
            // Resize([RowSize], [ColumnSize])
            // Returns a Range resized to the specified dimensions
            // Starts from the top-left cell of the original range
            let row_size = args.get(0).map(value_to_int).unwrap_or(1) as i32;
            let col_size = args.get(1).map(value_to_int).unwrap_or(1) as i32;
            
            if row_size < 1 || col_size < 1 {
                bail!("Resize dimensions must be >= 1");
            }
            
            // Get the top-left corner of the range
            let ((start_row, start_col), _) = get_range_bounds(address)?;
            
            // Calculate new end position
            let new_end_row = start_row + row_size - 1;
            let new_end_col = start_col + col_size - 1;
            
            // Create the new address
            let new_address = if row_size == 1 && col_size == 1 {
                indices_to_address(start_row, start_col)
            } else {
                format!("{}:{}", 
                    indices_to_address(start_row, start_col),
                    indices_to_address(new_end_row, new_end_col))
            };
            
            Ok(Value::String(format!("Range:{}", new_address)))
        }
        
        // ====================================================================
        // AUTOFIT
        // ====================================================================
        
        "autofit" => {
            // AutoFit for Columns or Rows (depends on which is called)
            // Usually Range.Columns.AutoFit or Range.Rows.AutoFit
            // TODO: ENGINE CALL - engine::autofit(address)
            eprintln!("   [STUB] Range({}).AutoFit()", address);
            Ok(Value::Empty)
        }
        
        // ====================================================================
        // UNKNOWN METHOD
        // ====================================================================
        
        _ => bail!("Unknown Range method: {}", method),
    }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

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
    }
    
    #[test]
    fn test_value_helpers() {
        assert_eq!(value_to_bool(&Value::Boolean(true)), true);
        assert_eq!(value_to_bool(&Value::Integer(0)), false);
        assert_eq!(value_to_int(&Value::Double(3.7)), 3);
        assert_eq!(value_to_string(&Value::Integer(42)), "42");
    }
}
