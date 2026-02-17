// src/host/excel/static_engine.rs
// ============================================================================
// STATIC ENGINE - Default/Stub Implementations for Excel Engine Functions
//
// This file provides default implementations for all engine functions needed
// by the VBA Range object. These are placeholder implementations that can be
// used when the native engine is not available or for testing purposes.
//
// HOW TO USE:
// 1. For each function, the native engine should implement the equivalent
//    C function with the same parameters
// 2. The functions here serve as documentation and fallback
// 3. In production, these are replaced by calls to NativeClientEngine library
//
// FUNCTION NAMING CONVENTION:
// - static_* : Functions that work without native engine (pure Rust)
// - engine_* : Functions that require native engine integration
//
// ============================================================================

use std::collections::HashMap;
use std::sync::Mutex;
use once_cell::sync::Lazy;

// ============================================================================
// IN-MEMORY STORAGE (for testing/stub mode)
// ============================================================================

/// In-memory cell storage for stub mode
/// Key: "SheetName!Row:Col" (0-based indices)
static CELL_STORAGE: Lazy<Mutex<HashMap<String, CellData>>> = Lazy::new(|| {
    Mutex::new(HashMap::new())
});

/// In-memory format storage
static FORMAT_STORAGE: Lazy<Mutex<HashMap<String, CellFormat>>> = Lazy::new(|| {
    Mutex::new(HashMap::new())
});

/// In-memory comment storage
static COMMENT_STORAGE: Lazy<Mutex<HashMap<String, String>>> = Lazy::new(|| {
    Mutex::new(HashMap::new())
});

/// In-memory merge storage (stores top-left cell of merge region)
static MERGE_STORAGE: Lazy<Mutex<HashMap<String, String>>> = Lazy::new(|| {
    Mutex::new(HashMap::new())
});

/// Cell data structure
#[derive(Clone, Debug, Default)]
pub struct CellData {
    pub value: String,
    pub formula: Option<String>,
    pub formula_r1c1: Option<String>,
    pub is_array_formula: bool,
}

/// Cell format structure  
#[derive(Clone, Debug)]
pub struct CellFormat {
    pub number_format: String,
    pub horizontal_alignment: i32,  // xlGeneral=-4105, xlLeft=-4131, xlCenter=-4108, xlRight=-4152
    pub vertical_alignment: i32,    // xlTop=-4160, xlCenter=-4108, xlBottom=-4107
    pub orientation: i32,           // -90 to 90 degrees
    pub wrap_text: bool,
    pub add_indent: bool,
    pub indent_level: i32,          // 0-15
    pub locked: bool,
    pub hidden: bool,
    pub font: FontFormat,
    pub interior: InteriorFormat,
    pub borders: BordersFormat,
}

impl Default for CellFormat {
    fn default() -> Self {
        Self {
            number_format: "General".to_string(),
            horizontal_alignment: -4105, // xlGeneral
            vertical_alignment: -4107,   // xlBottom
            orientation: 0,
            wrap_text: false,
            add_indent: false,
            indent_level: 0,
            locked: true,   // Default is locked
            hidden: false,
            font: FontFormat::default(),
            interior: InteriorFormat::default(),
            borders: BordersFormat::default(),
        }
    }
}

/// Font format structure
#[derive(Clone, Debug)]
pub struct FontFormat {
    pub name: String,
    pub size: f64,
    pub bold: bool,
    pub italic: bool,
    pub underline: i32,        // xlUnderlineStyleNone=0, xlUnderlineStyleSingle=-4142, etc.
    pub strikethrough: bool,
    pub color: i64,            // RGB color as Long
    pub color_index: i32,      // xlColorIndexAutomatic=-4105
}

impl Default for FontFormat {
    fn default() -> Self {
        Self {
            name: "Calibri".to_string(),
            size: 11.0,
            bold: false,
            italic: false,
            underline: -4142, // xlUnderlineStyleNone
            strikethrough: false,
            color: 0,         // Black
            color_index: -4105, // xlColorIndexAutomatic
        }
    }
}

/// Interior (fill) format structure
#[derive(Clone, Debug)]
pub struct InteriorFormat {
    pub color: i64,            // RGB color as Long
    pub color_index: i32,      // xlColorIndexNone=-4142
    pub pattern: i32,          // xlPatternNone=-4142, xlPatternSolid=1
    pub pattern_color: i64,
    pub pattern_color_index: i32,
}

impl Default for InteriorFormat {
    fn default() -> Self {
        Self {
            color: 16777215,      // White
            color_index: -4142,   // xlColorIndexNone (no fill)
            pattern: -4142,       // xlPatternNone
            pattern_color: 0,
            pattern_color_index: -4105,
        }
    }
}

/// Borders format structure
#[derive(Clone, Debug, Default)]
pub struct BordersFormat {
    pub left: BorderFormat,
    pub right: BorderFormat,
    pub top: BorderFormat,
    pub bottom: BorderFormat,
    pub diagonal_down: BorderFormat,
    pub diagonal_up: BorderFormat,
}

/// Single border format
#[derive(Clone, Debug)]
pub struct BorderFormat {
    pub line_style: i32,       // xlLineStyleNone=-4142, xlContinuous=1, etc.
    pub weight: i32,           // xlThin=2, xlMedium=-4138, xlThick=4
    pub color: i64,
    pub color_index: i32,
}

impl Default for BorderFormat {
    fn default() -> Self {
        Self {
            line_style: -4142, // xlLineStyleNone
            weight: 2,         // xlThin
            color: 0,
            color_index: -4105,
        }
    }
}

// ============================================================================
// CELL VALUE FUNCTIONS
// ============================================================================

/// Get cell value (static implementation)
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `row`: i32 - 0-based row index
/// - `col`: i32 - 0-based column index
/// 
/// # Returns
/// - String - Cell value as string
pub fn static_get_cell_value(sheet_name: &str, row: i32, col: i32) -> String {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = CELL_STORAGE.lock().unwrap();
    storage.get(&key)
        .map(|d| d.value.clone())
        .unwrap_or_default()
}

/// Set cell value (static implementation)
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `row`: i32 - 0-based row index
/// - `col`: i32 - 0-based column index
/// - `value`: &str - Value to set
/// 
/// # Returns
/// - bool - Success
pub fn static_set_cell_value(sheet_name: &str, row: i32, col: i32, value: &str) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = CELL_STORAGE.lock().unwrap();
    let entry = storage.entry(key).or_insert_with(CellData::default);
    entry.value = value.to_string();
    true
}

/// Get cell formatted text (as displayed)
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `row`: i32 - 0-based row index
/// - `col`: i32 - 0-based column index
/// 
/// # Returns
/// - String - Formatted text
pub fn static_get_cell_text(sheet_name: &str, row: i32, col: i32) -> String {
    // In static mode, just return the value
    // Real engine would apply number format
    static_get_cell_value(sheet_name, row, col)
}

// ============================================================================
// FORMULA FUNCTIONS
// ============================================================================

/// Get cell formula in A1 notation
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name  
/// - `row`: i32 - 0-based row index
/// - `col`: i32 - 0-based column index
/// 
/// # Returns
/// - String - Formula (empty if no formula)
pub fn static_get_cell_formula(sheet_name: &str, row: i32, col: i32) -> String {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = CELL_STORAGE.lock().unwrap();
    storage.get(&key)
        .and_then(|d| d.formula.clone())
        .unwrap_or_default()
}

/// Set cell formula in A1 notation
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `row`: i32 - 0-based row index
/// - `col`: i32 - 0-based column index
/// - `formula`: &str - Formula (must start with =)
/// 
/// # Returns
/// - bool - Success
pub fn static_set_cell_formula(sheet_name: &str, row: i32, col: i32, formula: &str) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = CELL_STORAGE.lock().unwrap();
    let entry = storage.entry(key).or_insert_with(CellData::default);
    entry.formula = Some(formula.to_string());
    // In real engine, this would trigger recalculation
    true
}

/// Get cell formula in R1C1 notation
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `row`: i32 - 0-based row index
/// - `col`: i32 - 0-based column index
/// 
/// # Returns
/// - String - Formula in R1C1 notation
pub fn static_get_cell_formula_r1c1(sheet_name: &str, row: i32, col: i32) -> String {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = CELL_STORAGE.lock().unwrap();
    storage.get(&key)
        .and_then(|d| d.formula_r1c1.clone())
        .unwrap_or_default()
}

/// Set cell formula in R1C1 notation
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `row`: i32 - 0-based row index
/// - `col`: i32 - 0-based column index
/// - `formula`: &str - Formula in R1C1 notation
/// 
/// # Returns
/// - bool - Success
pub fn static_set_cell_formula_r1c1(sheet_name: &str, row: i32, col: i32, formula: &str) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = CELL_STORAGE.lock().unwrap();
    let entry = storage.entry(key).or_insert_with(CellData::default);
    entry.formula_r1c1 = Some(formula.to_string());
    true
}

/// Get array formula for range
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `start_row`: i32 - 0-based start row
/// - `start_col`: i32 - 0-based start column
/// - `end_row`: i32 - 0-based end row
/// - `end_col`: i32 - 0-based end column
/// 
/// # Returns
/// - String - Array formula
pub fn static_get_array_formula(sheet_name: &str, start_row: i32, start_col: i32, _end_row: i32, _end_col: i32) -> String {
    // For array formula, we store it in the top-left cell
    static_get_cell_formula(sheet_name, start_row, start_col)
}

/// Set array formula for range
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `start_row`: i32 - 0-based start row
/// - `start_col`: i32 - 0-based start column
/// - `end_row`: i32 - 0-based end row
/// - `end_col`: i32 - 0-based end column
/// - `formula`: &str - Array formula
/// 
/// # Returns
/// - bool - Success
pub fn static_set_array_formula(sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32, formula: &str) -> bool {
    // Mark all cells as part of array formula
    for row in start_row..=end_row {
        for col in start_col..=end_col {
            let key = format!("{}!{}:{}", sheet_name, row, col);
            let mut storage = CELL_STORAGE.lock().unwrap();
            let entry = storage.entry(key).or_insert_with(CellData::default);
            entry.is_array_formula = true;
            if row == start_row && col == start_col {
                entry.formula = Some(formula.to_string());
            }
        }
    }
    true
}

/// Check if cell is part of array formula
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `row`: i32 - 0-based row index
/// - `col`: i32 - 0-based column index
/// 
/// # Returns
/// - bool - True if part of array formula
pub fn static_has_array_formula(sheet_name: &str, row: i32, col: i32) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = CELL_STORAGE.lock().unwrap();
    storage.get(&key)
        .map(|d| d.is_array_formula)
        .unwrap_or(false)
}

// ============================================================================
// NUMBER FORMAT FUNCTIONS
// ============================================================================

/// Get cell number format
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `row`: i32 - 0-based row index
/// - `col`: i32 - 0-based column index
/// 
/// # Returns
/// - String - Number format code (e.g., "General", "0.00", "@")
pub fn static_get_number_format(sheet_name: &str, row: i32, col: i32) -> String {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = FORMAT_STORAGE.lock().unwrap();
    storage.get(&key)
        .map(|f| f.number_format.clone())
        .unwrap_or_else(|| "General".to_string())
}

/// Set cell number format
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `row`: i32 - 0-based row index
/// - `col`: i32 - 0-based column index
/// - `format`: &str - Number format code
/// 
/// # Returns
/// - bool - Success
pub fn static_set_number_format(sheet_name: &str, row: i32, col: i32, format: &str) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = FORMAT_STORAGE.lock().unwrap();
    let entry = storage.entry(key).or_insert_with(CellFormat::default);
    entry.number_format = format.to_string();
    true
}

// ============================================================================
// ALIGNMENT FUNCTIONS
// ============================================================================

/// Get horizontal alignment
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `row`: i32 - 0-based row index
/// - `col`: i32 - 0-based column index
/// 
/// # Returns
/// - i32 - Alignment constant (xlGeneral=-4105, xlLeft=-4131, xlCenter=-4108, xlRight=-4152)
pub fn static_get_horizontal_alignment(sheet_name: &str, row: i32, col: i32) -> i32 {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = FORMAT_STORAGE.lock().unwrap();
    storage.get(&key)
        .map(|f| f.horizontal_alignment)
        .unwrap_or(-4105) // xlGeneral
}

/// Set horizontal alignment
pub fn static_set_horizontal_alignment(sheet_name: &str, row: i32, col: i32, alignment: i32) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = FORMAT_STORAGE.lock().unwrap();
    let entry = storage.entry(key).or_insert_with(CellFormat::default);
    entry.horizontal_alignment = alignment;
    true
}

/// Get vertical alignment
pub fn static_get_vertical_alignment(sheet_name: &str, row: i32, col: i32) -> i32 {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = FORMAT_STORAGE.lock().unwrap();
    storage.get(&key)
        .map(|f| f.vertical_alignment)
        .unwrap_or(-4107) // xlBottom
}

/// Set vertical alignment
pub fn static_set_vertical_alignment(sheet_name: &str, row: i32, col: i32, alignment: i32) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = FORMAT_STORAGE.lock().unwrap();
    let entry = storage.entry(key).or_insert_with(CellFormat::default);
    entry.vertical_alignment = alignment;
    true
}

/// Get text orientation (-90 to 90 degrees)
pub fn static_get_orientation(sheet_name: &str, row: i32, col: i32) -> i32 {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = FORMAT_STORAGE.lock().unwrap();
    storage.get(&key)
        .map(|f| f.orientation)
        .unwrap_or(0)
}

/// Set text orientation
pub fn static_set_orientation(sheet_name: &str, row: i32, col: i32, degrees: i32) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = FORMAT_STORAGE.lock().unwrap();
    let entry = storage.entry(key).or_insert_with(CellFormat::default);
    entry.orientation = degrees.clamp(-90, 90);
    true
}

/// Get wrap text setting
pub fn static_get_wrap_text(sheet_name: &str, row: i32, col: i32) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = FORMAT_STORAGE.lock().unwrap();
    storage.get(&key)
        .map(|f| f.wrap_text)
        .unwrap_or(false)
}

/// Set wrap text setting
pub fn static_set_wrap_text(sheet_name: &str, row: i32, col: i32, wrap: bool) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = FORMAT_STORAGE.lock().unwrap();
    let entry = storage.entry(key).or_insert_with(CellFormat::default);
    entry.wrap_text = wrap;
    true
}

/// Get indent level (0-15)
pub fn static_get_indent_level(sheet_name: &str, row: i32, col: i32) -> i32 {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = FORMAT_STORAGE.lock().unwrap();
    storage.get(&key)
        .map(|f| f.indent_level)
        .unwrap_or(0)
}

/// Set indent level
pub fn static_set_indent_level(sheet_name: &str, row: i32, col: i32, level: i32) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = FORMAT_STORAGE.lock().unwrap();
    let entry = storage.entry(key).or_insert_with(CellFormat::default);
    entry.indent_level = level.clamp(0, 15);
    true
}

// ============================================================================
// CELL STATE FUNCTIONS
// ============================================================================

/// Get locked state
pub fn static_get_locked(sheet_name: &str, row: i32, col: i32) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = FORMAT_STORAGE.lock().unwrap();
    storage.get(&key)
        .map(|f| f.locked)
        .unwrap_or(true) // Default is locked
}

/// Set locked state
pub fn static_set_locked(sheet_name: &str, row: i32, col: i32, locked: bool) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = FORMAT_STORAGE.lock().unwrap();
    let entry = storage.entry(key).or_insert_with(CellFormat::default);
    entry.locked = locked;
    true
}

/// Get hidden state
pub fn static_get_hidden(sheet_name: &str, row: i32, col: i32) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = FORMAT_STORAGE.lock().unwrap();
    storage.get(&key)
        .map(|f| f.hidden)
        .unwrap_or(false)
}

/// Set hidden state
pub fn static_set_hidden(sheet_name: &str, row: i32, col: i32, hidden: bool) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = FORMAT_STORAGE.lock().unwrap();
    let entry = storage.entry(key).or_insert_with(CellFormat::default);
    entry.hidden = hidden;
    true
}

// ============================================================================
// MERGE CELL FUNCTIONS
// ============================================================================

/// Check if cell is part of merged region
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `row`: i32 - 0-based row index
/// - `col`: i32 - 0-based column index
/// 
/// # Returns
/// - bool - True if merged
pub fn static_is_merged(sheet_name: &str, row: i32, col: i32) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = MERGE_STORAGE.lock().unwrap();
    storage.contains_key(&key)
}

/// Merge cells
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `start_row`: i32 - 0-based start row
/// - `start_col`: i32 - 0-based start column
/// - `end_row`: i32 - 0-based end row
/// - `end_col`: i32 - 0-based end column
/// - `across`: bool - If true, merge each row separately
/// 
/// # Returns
/// - bool - Success
pub fn static_merge_cells(sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32, across: bool) -> bool {
    let mut storage = MERGE_STORAGE.lock().unwrap();
    let top_left = format!("{}:{}", start_row, start_col);
    
    if across {
        // Merge each row separately
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let key = format!("{}!{}:{}", sheet_name, row, col);
                storage.insert(key, format!("{}:{}", row, start_col));
            }
        }
    } else {
        // Merge entire range
        for row in start_row..=end_row {
            for col in start_col..=end_col {
                let key = format!("{}!{}:{}", sheet_name, row, col);
                storage.insert(key, top_left.clone());
            }
        }
    }
    true
}

/// Unmerge cells
pub fn static_unmerge_cells(sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> bool {
    let mut storage = MERGE_STORAGE.lock().unwrap();
    for row in start_row..=end_row {
        for col in start_col..=end_col {
            let key = format!("{}!{}:{}", sheet_name, row, col);
            storage.remove(&key);
        }
    }
    true
}

// ============================================================================
// COMMENT FUNCTIONS
// ============================================================================

/// Get cell comment
pub fn static_get_comment(sheet_name: &str, row: i32, col: i32) -> Option<String> {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let storage = COMMENT_STORAGE.lock().unwrap();
    storage.get(&key).cloned()
}

/// Add cell comment
pub fn static_add_comment(sheet_name: &str, row: i32, col: i32, text: &str) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = COMMENT_STORAGE.lock().unwrap();
    storage.insert(key, text.to_string());
    true
}

/// Clear cell comment
pub fn static_clear_comment(sheet_name: &str, row: i32, col: i32) -> bool {
    let key = format!("{}!{}:{}", sheet_name, row, col);
    let mut storage = COMMENT_STORAGE.lock().unwrap();
    storage.remove(&key);
    true
}

// ============================================================================
// SELECTION & ACTIVATION FUNCTIONS
// ============================================================================

/// Select range (UI operation - stub in static mode)
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `start_row`: i32 - 0-based start row
/// - `start_col`: i32 - 0-based start column  
/// - `end_row`: i32 - 0-based end row
/// - `end_col`: i32 - 0-based end column
/// 
/// # Returns
/// - bool - Success
pub fn static_select_range(_sheet_name: &str, _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32) -> bool {
    // In static mode, this is a no-op
    // Real engine would update UI selection
    true
}

/// Activate cell
pub fn static_activate_cell(_sheet_name: &str, _row: i32, _col: i32) -> bool {
    true
}

// ============================================================================
// CLIPBOARD FUNCTIONS
// ============================================================================

/// Copy range to clipboard
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `start_row`: i32 - 0-based start row
/// - `start_col`: i32 - 0-based start column
/// - `end_row`: i32 - 0-based end row
/// - `end_col`: i32 - 0-based end column
/// 
/// # Returns
/// - bool - Success
pub fn static_copy_range(_sheet_name: &str, _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32) -> bool {
    // Would copy to internal clipboard
    true
}

/// Copy range to destination
/// 
/// # Parameters
/// - `src_sheet`: &str - Source sheet name
/// - `src_start_row`: i32 - Source start row
/// - `src_start_col`: i32 - Source start column
/// - `src_end_row`: i32 - Source end row
/// - `src_end_col`: i32 - Source end column
/// - `dest_sheet`: &str - Destination sheet name
/// - `dest_row`: i32 - Destination top-left row
/// - `dest_col`: i32 - Destination top-left column
/// 
/// # Returns
/// - bool - Success
pub fn static_copy_range_to(
    src_sheet: &str, src_start_row: i32, src_start_col: i32, src_end_row: i32, src_end_col: i32,
    dest_sheet: &str, dest_row: i32, dest_col: i32
) -> bool {
    let row_count = src_end_row - src_start_row;
    let col_count = src_end_col - src_start_col;
    
    for r in 0..=row_count {
        for c in 0..=col_count {
            let value = static_get_cell_value(src_sheet, src_start_row + r, src_start_col + c);
            static_set_cell_value(dest_sheet, dest_row + r, dest_col + c, &value);
        }
    }
    true
}

/// Cut range to clipboard
pub fn static_cut_range(_sheet_name: &str, _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32) -> bool {
    true
}

/// Paste special
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `row`: i32 - Destination row
/// - `col`: i32 - Destination column
/// - `paste_type`: i32 - xlPasteAll(-4104), xlPasteValues(-4163), etc.
/// - `operation`: i32 - xlNone(-4142), xlAdd(2), xlSubtract(3), etc.
/// - `skip_blanks`: bool - Skip blank cells
/// - `transpose`: bool - Transpose rows/columns
/// 
/// # Returns
/// - bool - Success
pub fn static_paste_special(
    _sheet_name: &str, _row: i32, _col: i32,
    _paste_type: i32, _operation: i32, _skip_blanks: bool, _transpose: bool
) -> bool {
    true
}

// ============================================================================
// CLEAR FUNCTIONS
// ============================================================================

/// Clear range (all: values, formats, comments)
pub fn static_clear_range(sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> bool {
    for row in start_row..=end_row {
        for col in start_col..=end_col {
            static_set_cell_value(sheet_name, row, col, "");
            let key = format!("{}!{}:{}", sheet_name, row, col);
            FORMAT_STORAGE.lock().unwrap().remove(&key);
            COMMENT_STORAGE.lock().unwrap().remove(&key);
        }
    }
    true
}

/// Clear contents only (values and formulas)
pub fn static_clear_contents(sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> bool {
    for row in start_row..=end_row {
        for col in start_col..=end_col {
            static_set_cell_value(sheet_name, row, col, "");
        }
    }
    true
}

/// Clear formats only
pub fn static_clear_formats(sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> bool {
    for row in start_row..=end_row {
        for col in start_col..=end_col {
            let key = format!("{}!{}:{}", sheet_name, row, col);
            FORMAT_STORAGE.lock().unwrap().remove(&key);
        }
    }
    true
}

/// Clear comments only
pub fn static_clear_comments(sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> bool {
    for row in start_row..=end_row {
        for col in start_col..=end_col {
            static_clear_comment(sheet_name, row, col);
        }
    }
    true
}

// ============================================================================
// FIND & REPLACE FUNCTIONS
// ============================================================================

/// Find value in range
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `start_row`: i32 - Search start row
/// - `start_col`: i32 - Search start column
/// - `end_row`: i32 - Search end row
/// - `end_col`: i32 - Search end column
/// - `what`: &str - Value to find
/// - `look_in`: i32 - xlValues(-4163), xlFormulas(-4123), xlComments(-4144)
/// - `look_at`: i32 - xlWhole(1), xlPart(2)
/// - `match_case`: bool - Case sensitive
/// 
/// # Returns
/// - Option<(i32, i32)> - (row, col) of first match, or None
pub fn static_find_in_range(
    sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32,
    what: &str, _look_in: i32, look_at: i32, match_case: bool
) -> Option<(i32, i32)> {
    let search = if match_case { what.to_string() } else { what.to_lowercase() };
    
    for row in start_row..=end_row {
        for col in start_col..=end_col {
            let value = static_get_cell_value(sheet_name, row, col);
            let check = if match_case { value.clone() } else { value.to_lowercase() };
            
            let found = if look_at == 1 { // xlWhole
                check == search
            } else { // xlPart
                check.contains(&search)
            };
            
            if found {
                return Some((row, col));
            }
        }
    }
    None
}

/// Replace value in range
/// 
/// # Returns
/// - i32 - Number of replacements made
pub fn static_replace_in_range(
    sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32,
    what: &str, replacement: &str, look_at: i32, match_case: bool
) -> i32 {
    let mut count = 0;
    
    for row in start_row..=end_row {
        for col in start_col..=end_col {
            let value = static_get_cell_value(sheet_name, row, col);
            let new_value = if look_at == 1 { // xlWhole
                let matches = if match_case {
                    value == what
                } else {
                    value.eq_ignore_ascii_case(what)
                };
                if matches {
                    replacement.to_string()
                } else {
                    continue;
                }
            } else { // xlPart
                if match_case {
                    if value.contains(what) {
                        value.replace(what, replacement)
                    } else {
                        continue;
                    }
                } else {
                    // Case-insensitive replace is more complex
                    let lower = value.to_lowercase();
                    let what_lower = what.to_lowercase();
                    if lower.contains(&what_lower) {
                        // Simple approach: just do case-sensitive replace
                        value.replace(what, replacement)
                    } else {
                        continue;
                    }
                }
            };
            
            static_set_cell_value(sheet_name, row, col, &new_value);
            count += 1;
        }
    }
    count
}

// ============================================================================
// INSERT & DELETE FUNCTIONS
// ============================================================================

/// Insert cells
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `start_row`: i32 - Start row
/// - `start_col`: i32 - Start column
/// - `end_row`: i32 - End row
/// - `end_col`: i32 - End column
/// - `shift`: i32 - xlShiftDown(-4121) or xlShiftToRight(-4161)
/// 
/// # Returns
/// - bool - Success
pub fn static_insert_cells(
    _sheet_name: &str, _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32,
    _shift: i32
) -> bool {
    // Complex operation requiring shifting all cells
    // Would be implemented in native engine
    true
}

/// Delete cells
pub fn static_delete_cells(
    _sheet_name: &str, _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32,
    _shift: i32
) -> bool {
    true
}

// ============================================================================
// FILL FUNCTIONS
// ============================================================================

/// Auto fill from source range to destination
pub fn static_auto_fill(
    _sheet_name: &str, 
    _src_start_row: i32, _src_start_col: i32, _src_end_row: i32, _src_end_col: i32,
    _dest_start_row: i32, _dest_start_col: i32, _dest_end_row: i32, _dest_end_col: i32,
    _fill_type: i32
) -> bool {
    true
}

/// Fill down
pub fn static_fill_down(sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> bool {
    for col in start_col..=end_col {
        let value = static_get_cell_value(sheet_name, start_row, col);
        for row in (start_row + 1)..=end_row {
            static_set_cell_value(sheet_name, row, col, &value);
        }
    }
    true
}

/// Fill up
pub fn static_fill_up(sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> bool {
    for col in start_col..=end_col {
        let value = static_get_cell_value(sheet_name, end_row, col);
        for row in start_row..end_row {
            static_set_cell_value(sheet_name, row, col, &value);
        }
    }
    true
}

/// Fill left
pub fn static_fill_left(sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> bool {
    for row in start_row..=end_row {
        let value = static_get_cell_value(sheet_name, row, end_col);
        for col in start_col..end_col {
            static_set_cell_value(sheet_name, row, col, &value);
        }
    }
    true
}

/// Fill right
pub fn static_fill_right(sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32) -> bool {
    for row in start_row..=end_row {
        let value = static_get_cell_value(sheet_name, row, start_col);
        for col in (start_col + 1)..=end_col {
            static_set_cell_value(sheet_name, row, col, &value);
        }
    }
    true
}

// ============================================================================
// SORT & FILTER FUNCTIONS
// ============================================================================

/// Sort range
pub fn static_sort_range(
    _sheet_name: &str, _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32,
    _key1_col: i32, _order1: i32, _has_header: bool
) -> bool {
    // Complex operation - would be in native engine
    true
}

/// Apply auto filter
pub fn static_auto_filter(
    _sheet_name: &str, _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32,
    _field: Option<i32>, _criteria1: Option<&str>
) -> bool {
    true
}

/// Remove duplicates
pub fn static_remove_duplicates(
    _sheet_name: &str, _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32,
    _columns: &[i32], _has_header: bool
) -> bool {
    true
}

// ============================================================================
// CALCULATION FUNCTIONS
// ============================================================================

/// Calculate range
pub fn static_calculate_range(_sheet_name: &str, _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32) -> bool {
    // Would trigger formula recalculation
    true
}

// ============================================================================
// DEPENDENCY FUNCTIONS
// ============================================================================

/// Get direct dependents (cells that reference this cell)
pub fn static_get_direct_dependents(_sheet_name: &str, _row: i32, _col: i32) -> Vec<(i32, i32)> {
    Vec::new()
}

/// Get direct precedents (cells that this cell references)
pub fn static_get_direct_precedents(_sheet_name: &str, _row: i32, _col: i32) -> Vec<(i32, i32)> {
    Vec::new()
}

/// Get all dependents (recursive)
pub fn static_get_dependents(_sheet_name: &str, _row: i32, _col: i32) -> Vec<(i32, i32)> {
    Vec::new()
}

/// Get all precedents (recursive)
pub fn static_get_precedents(_sheet_name: &str, _row: i32, _col: i32) -> Vec<(i32, i32)> {
    Vec::new()
}

// ============================================================================
// SPECIAL CELLS FUNCTION
// ============================================================================

/// Get special cells matching criteria
/// 
/// # Parameters
/// - `sheet_name`: &str - Sheet name
/// - `start_row`: i32 - Start row
/// - `start_col`: i32 - Start column
/// - `end_row`: i32 - End row
/// - `end_col`: i32 - End column
/// - `cell_type`: i32 - xlCellTypeConstants(2), xlCellTypeFormulas(-4123), xlCellTypeBlanks(4), etc.
/// - `value_type`: Option<i32> - xlNumbers(1), xlTextValues(2), xlLogical(4), xlErrors(16)
/// 
/// # Returns
/// - Vec<(i32, i32)> - List of matching cell coordinates
pub fn static_get_special_cells(
    sheet_name: &str, start_row: i32, start_col: i32, end_row: i32, end_col: i32,
    cell_type: i32, _value_type: Option<i32>
) -> Vec<(i32, i32)> {
    let mut results = Vec::new();
    
    for row in start_row..=end_row {
        for col in start_col..=end_col {
            let value = static_get_cell_value(sheet_name, row, col);
            let formula = static_get_cell_formula(sheet_name, row, col);
            
            let matches = match cell_type {
                4 => value.is_empty() && formula.is_empty(), // xlCellTypeBlanks
                2 => !value.is_empty() && formula.is_empty(), // xlCellTypeConstants
                -4123 => !formula.is_empty(), // xlCellTypeFormulas
                _ => false,
            };
            
            if matches {
                results.push((row, col));
            }
        }
    }
    results
}

// ============================================================================
// CURRENT REGION FUNCTION
// ============================================================================

/// Get current region (contiguous non-empty cells)
/// 
/// # Returns
/// - (start_row, start_col, end_row, end_col) bounds of current region
pub fn static_get_current_region(sheet_name: &str, row: i32, col: i32) -> (i32, i32, i32, i32) {
    // Simplified: just return the cell itself
    // Real implementation would expand to find contiguous data
    let _ = (sheet_name, row, col);
    (row, col, row, col)
}

// ============================================================================
// STYLE FUNCTIONS
// ============================================================================

/// Get cell style name
pub fn static_get_style(_sheet_name: &str, _row: i32, _col: i32) -> String {
    "Normal".to_string()
}

/// Set cell style
pub fn static_set_style(_sheet_name: &str, _row: i32, _col: i32, _style_name: &str) -> bool {
    true
}

// ============================================================================
// NAMED RANGE FUNCTIONS
// ============================================================================

/// Get name for range
pub fn static_get_range_name(_sheet_name: &str, _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32) -> Option<String> {
    None
}

/// Create named range
pub fn static_create_named_range(
    _name: &str, _sheet_name: &str, 
    _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32
) -> bool {
    true
}

// ============================================================================
// HYPERLINK FUNCTIONS
// ============================================================================

/// Get hyperlink from cell
pub fn static_get_hyperlink(_sheet_name: &str, _row: i32, _col: i32) -> Option<String> {
    None
}

/// Add hyperlink to cell
pub fn static_add_hyperlink(_sheet_name: &str, _row: i32, _col: i32, _address: &str, _text: &str) -> bool {
    true
}

// ============================================================================
// VALIDATION FUNCTIONS
// ============================================================================

/// Get data validation for cell
pub fn static_get_validation(_sheet_name: &str, _row: i32, _col: i32) -> Option<ValidationInfo> {
    None
}

/// Validation info structure
pub struct ValidationInfo {
    pub validation_type: i32,
    pub formula1: String,
    pub formula2: Option<String>,
    pub operator: i32,
    pub alert_style: i32,
    pub input_title: String,
    pub input_message: String,
    pub error_title: String,
    pub error_message: String,
}

/// Set data validation
pub fn static_set_validation(
    _sheet_name: &str, _row: i32, _col: i32,
    _validation_type: i32, _formula1: &str, _formula2: Option<&str>, _operator: i32
) -> bool {
    true
}

// ============================================================================
// GROUP/OUTLINE FUNCTIONS
// ============================================================================

/// Group rows/columns
pub fn static_group(_sheet_name: &str, _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32) -> bool {
    true
}

/// Ungroup rows/columns
pub fn static_ungroup(_sheet_name: &str, _start_row: i32, _start_col: i32, _end_row: i32, _end_col: i32) -> bool {
    true
}

// ============================================================================
// AUTOFIT FUNCTIONS
// ============================================================================

/// Autofit columns
pub fn static_autofit_columns(_sheet_name: &str, _start_col: i32, _end_col: i32) -> bool {
    true
}

/// Autofit rows
pub fn static_autofit_rows(_sheet_name: &str, _start_row: i32, _end_row: i32) -> bool {
    true
}

// ============================================================================
// TESTS
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_cell_value() {
        static_set_cell_value("Sheet1", 0, 0, "Hello");
        assert_eq!(static_get_cell_value("Sheet1", 0, 0), "Hello");
    }

    #[test]
    fn test_number_format() {
        static_set_number_format("Sheet1", 0, 0, "0.00");
        assert_eq!(static_get_number_format("Sheet1", 0, 0), "0.00");
    }

    #[test]
    fn test_fill_down() {
        static_set_cell_value("Sheet1", 0, 0, "Test");
        static_fill_down("Sheet1", 0, 0, 2, 0);
        assert_eq!(static_get_cell_value("Sheet1", 1, 0), "Test");
        assert_eq!(static_get_cell_value("Sheet1", 2, 0), "Test");
    }
}
