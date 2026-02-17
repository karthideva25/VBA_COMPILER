// src/host/excel/objects/range.rs
// ============================================================================
// Excel Range Object - COM-style implementation
//
// The Range object is the most important object in Excel VBA automation.
// It represents a cell, a selection of cells, or a 3D range spanning multiple sheets.
//
// Architecture:
// - Range stores its address (e.g., "A1", "A1:B5", "Sheet1!A1:B5")
// - Properties and methods are dispatched via the ComObject trait
// - The engine.rs handles actual cell value read/write to the native engine
//
// Usage patterns in VBA:
// - Range("A1").Value = 10
// - x = Range("A1:B5").Count
// - Range("A1").Select
// - Set rng = Range("A1:B5")
// - rng.Copy Destination:=Range("C1")
// ============================================================================

use std::collections::HashMap;
use anyhow::Result;
use crate::context::{Context, Value};
use crate::host::ComObject;
use crate::host::excel::{engine, properties, methods};

/// Excel Range Object
/// 
/// Represents a cell, a row, a column, a selection of cells containing 
/// one or more contiguous blocks of cells, or a 3-D range.
#[derive(Debug, Clone)]
pub struct ExcelRange {
    /// The address of the range (e.g., "A1", "A1:B5", "Sheet1!A1:B5")
    pub address: String,
    
    /// Optional sheet name (if range is on a specific sheet)
    pub sheet_name: Option<String>,
    
    /// Parent worksheet reference (for .Parent property)
    pub parent_sheet: Option<String>,
    
    /// Cached values for performance (optional)
    cache: RangeCache,
}

/// Internal cache for frequently accessed properties
#[derive(Debug, Clone, Default)]
struct RangeCache {
    /// Cached row count
    pub row_count: Option<i64>,
    /// Cached column count
    pub col_count: Option<i64>,
    /// Cached cell count
    pub cell_count: Option<i64>,
}

impl ExcelRange {
    /// Create a new Range object from an address
    pub fn new(address: impl Into<String>) -> Self {
        let addr = address.into();
        let (sheet, cell_addr) = Self::parse_address(&addr);
        
        Self {
            address: cell_addr,
            sheet_name: sheet,
            parent_sheet: None,
            cache: RangeCache::default(),
        }
    }
    
    /// Create a Range with explicit sheet name
    pub fn with_sheet(address: impl Into<String>, sheet: impl Into<String>) -> Self {
        Self {
            address: address.into(),
            sheet_name: Some(sheet.into()),
            parent_sheet: None,
            cache: RangeCache::default(),
        }
    }
    
    /// Parse address like "Sheet1!A1:B5" into (Some("Sheet1"), "A1:B5")
    fn parse_address(address: &str) -> (Option<String>, String) {
        if let Some(pos) = address.find('!') {
            let sheet = address[..pos].trim_matches('\'').to_string();
            let cell_addr = address[pos + 1..].to_string();
            (Some(sheet), cell_addr)
        } else {
            (None, address.to_string())
        }
    }
    
    /// Get the full address including sheet name if present
    pub fn full_address(&self) -> String {
        if let Some(ref sheet) = self.sheet_name {
            if sheet.contains(' ') {
                format!("'{}'!{}", sheet, self.address)
            } else {
                format!("{}!{}", sheet, self.address)
            }
        } else {
            self.address.clone()
        }
    }
    
    /// Check if this is a multi-cell range
    pub fn is_multi_cell(&self) -> bool {
        self.address.contains(':')
    }
    
    /// Get the start and end addresses for a range
    pub fn get_bounds(&self) -> Result<((i32, i32), (i32, i32))> {
        if self.address.contains(':') {
            let parts: Vec<&str> = self.address.split(':').collect();
            if parts.len() != 2 {
                anyhow::bail!("Invalid range address: {}", self.address);
            }
            let start = engine::address_to_indices(parts[0])
                .map_err(|e| anyhow::anyhow!("{}", e))?;
            let end = engine::address_to_indices(parts[1])
                .map_err(|e| anyhow::anyhow!("{}", e))?;
            Ok((start, end))
        } else {
            let pos = engine::address_to_indices(&self.address)
                .map_err(|e| anyhow::anyhow!("{}", e))?;
            Ok((pos, pos))
        }
    }
    
    /// Get the number of rows in the range
    pub fn row_count(&self) -> Result<i64> {
        let ((start_row, _), (end_row, _)) = self.get_bounds()?;
        Ok((end_row - start_row + 1).abs() as i64)
    }
    
    /// Get the number of columns in the range
    pub fn col_count(&self) -> Result<i64> {
        let ((_, start_col), (_, end_col)) = self.get_bounds()?;
        Ok((end_col - start_col + 1).abs() as i64)
    }
    
    /// Get the total number of cells in the range
    pub fn cell_count(&self) -> Result<i64> {
        Ok(self.row_count()? * self.col_count()?)
    }
    
    /// Get the top-left cell address
    pub fn top_left(&self) -> Result<String> {
        let ((row, col), _) = self.get_bounds()?;
        Ok(indices_to_address(row, col))
    }
    
    /// Get the bottom-right cell address
    pub fn bottom_right(&self) -> Result<String> {
        let (_, (row, col)) = self.get_bounds()?;
        Ok(indices_to_address(row, col))
    }
    
    /// Create an offset range
    pub fn offset(&self, row_offset: i32, col_offset: i32) -> Result<ExcelRange> {
        let ((start_row, start_col), (end_row, end_col)) = self.get_bounds()?;
        let new_start = indices_to_address(start_row + row_offset, start_col + col_offset);
        let new_end = indices_to_address(end_row + row_offset, end_col + col_offset);
        
        let new_addr = if self.is_multi_cell() {
            format!("{}:{}", new_start, new_end)
        } else {
            new_start
        };
        
        Ok(ExcelRange::new(new_addr))
    }
    
    /// Create a resized range
    pub fn resize(&self, new_rows: Option<i32>, new_cols: Option<i32>) -> Result<ExcelRange> {
        let ((start_row, start_col), (end_row, end_col)) = self.get_bounds()?;
        
        let rows = new_rows.unwrap_or((end_row - start_row + 1) as i32);
        let cols = new_cols.unwrap_or((end_col - start_col + 1) as i32);
        
        if rows < 1 || cols < 1 {
            anyhow::bail!("Resize dimensions must be >= 1");
        }
        
        let new_end_row = start_row + rows - 1;
        let new_end_col = start_col + cols - 1;
        
        let new_start = indices_to_address(start_row, start_col);
        let new_end = indices_to_address(new_end_row, new_end_col);
        
        let new_addr = if rows == 1 && cols == 1 {
            new_start
        } else {
            format!("{}:{}", new_start, new_end)
        };
        
        Ok(ExcelRange::new(new_addr))
    }
    
    /// Get a cell at specific row/column within the range (1-based)
    pub fn cells(&self, row: i32, col: i32) -> Result<ExcelRange> {
        let ((start_row, start_col), _) = self.get_bounds()?;
        let target_row = start_row + row - 1;
        let target_col = start_col + col - 1;
        
        if target_row < 0 || target_col < 0 {
            anyhow::bail!("Cell indices must be >= 1");
        }
        
        Ok(ExcelRange::new(indices_to_address(target_row, target_col)))
    }
}

/// Convert 0-based (row, col) to Excel address like "A1"
pub fn indices_to_address(row: i32, col: i32) -> String {
    let col_letter = column_index_to_letter(col);
    format!("{}{}", col_letter, row + 1)
}

/// Convert 0-based column index to Excel column letters (A, B, ..., Z, AA, AB, ...)
pub fn column_index_to_letter(col: i32) -> String {
    let mut result = String::new();
    let mut n = col + 1; // Convert to 1-based
    
    while n > 0 {
        n -= 1; // Adjust for 0-indexed calculation
        let remainder = (n % 26) as u8;
        result.insert(0, (b'A' + remainder) as char);
        n /= 26;
    }
    
    result
}

/// Implement ComObject trait for Range
impl ComObject for ExcelRange {
    fn get_property(&self, name: &str, ctx: &mut Context) -> Result<Value> {
        properties::range_properties::get_range_property(&self.address, name)
    }

    fn set_property(&mut self, name: &str, value: Value, ctx: &mut Context) -> Result<()> {
        properties::range_properties::set_range_property(&self.address, name, value)
    }

    fn call_method(&mut self, name: &str, args: &[Value], _ctx: &mut Context) -> Result<Value> {
        methods::range_methods::call_range_method(&self.address, name, args)
    }

    fn type_name(&self) -> &str {
        "Range"
    }
}

// ============================================================================
// RANGE BUILDER - Fluent API for creating ranges
// ============================================================================

/// Builder pattern for creating Range objects
pub struct RangeBuilder {
    start_row: i32,
    start_col: i32,
    end_row: Option<i32>,
    end_col: Option<i32>,
    sheet: Option<String>,
}

impl RangeBuilder {
    /// Start building a range from a cell
    pub fn from_cell(row: i32, col: i32) -> Self {
        Self {
            start_row: row,
            start_col: col,
            end_row: None,
            end_col: None,
            sheet: None,
        }
    }
    
    /// Set the end cell (makes it a multi-cell range)
    pub fn to_cell(mut self, row: i32, col: i32) -> Self {
        self.end_row = Some(row);
        self.end_col = Some(col);
        self
    }
    
    /// Set the sheet name
    pub fn on_sheet(mut self, sheet: impl Into<String>) -> Self {
        self.sheet = Some(sheet.into());
        self
    }
    
    /// Build the Range object
    pub fn build(self) -> ExcelRange {
        let start = indices_to_address(self.start_row, self.start_col);
        
        let address = if let (Some(end_row), Some(end_col)) = (self.end_row, self.end_col) {
            let end = indices_to_address(end_row, end_col);
            format!("{}:{}", start, end)
        } else {
            start
        };
        
        if let Some(sheet) = self.sheet {
            ExcelRange::with_sheet(address, sheet)
        } else {
            ExcelRange::new(address)
        }
    }
}
