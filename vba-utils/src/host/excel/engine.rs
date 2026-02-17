
// src/host/excel/engine.rs
use std::ffi::{CStr, CString, c_char, c_int};
use libc::free;
use std::sync::Mutex;
use once_cell::sync::Lazy;

#[link(name = "NativeClientEngine")]
unsafe extern "C" {
    fn EngineInterface_Initialize(resource_path: *const c_char, local_path: *const c_char) -> bool;
    fn EngineInterface_GetWorkbook() -> *const c_char;
    fn EngineInterface_GetSheet(workbook_id: *const c_char, sheet_name: *const c_char) -> *const c_char;
    fn EngineInterface_SetCellValue(workbook_id: *const c_char, sheet_name: *const c_char, row: c_int, col: c_int, value: *const c_char) -> bool;
    fn EngineInterface_GetCellValue(workbook_id: *const c_char, sheet_name: *const c_char, row: c_int, col: c_int) -> *const c_char;
}



/// Global state for the Excel engine
static ENGINE_STATE: Lazy<Mutex<EngineState>> = Lazy::new(|| {
    Mutex::new(EngineState {
        initialized: false,
        workbook_id: None,
        active_sheet: "Sheet1".to_string(),
    })
});

struct EngineState {
    initialized: bool,
    workbook_id: Option<String>,
    active_sheet: String,
}

/// Initialize the Excel engine
pub fn initialize_engine(resource_path: &str, local_path: &str) -> Result<(), String> {
    unsafe {
        let resource_cstr = CString::new(resource_path).map_err(|e| e.to_string())?;
        let local_cstr = CString::new(local_path).map_err(|e| e.to_string())?;
        
        let success = EngineInterface_Initialize(resource_cstr.as_ptr(), local_cstr.as_ptr());
        
        if success {
            // Get workbook ID
            let workbook_ptr = EngineInterface_GetWorkbook();
            if workbook_ptr.is_null() {
                return Err("Failed to get workbook ID".to_string());
            }
            
            let workbook_id = CStr::from_ptr(workbook_ptr)
                .to_str()
                .map_err(|e| e.to_string())?
                .to_owned();
            
            free(workbook_ptr as *mut libc::c_void);
            
            let mut state = ENGINE_STATE.lock().unwrap();
            state.initialized = true;
            state.workbook_id = Some(workbook_id.clone());
            
            // eprintln!("✅ Excel Engine initialized, Workbook ID: {}", workbook_id);
            Ok(())
        } else {
            Err("EngineInterface_Initialize failed".to_string())
        }
    }
}

/// Check if engine is initialized
pub fn is_initialized() -> bool {
    ENGINE_STATE.lock().unwrap().initialized
}

/// Get current workbook ID
pub fn get_workbook_id() -> Option<String> {
    ENGINE_STATE.lock().unwrap().workbook_id.clone()
}

/// Set active sheet
pub fn set_active_sheet(sheet_name: String) {
    ENGINE_STATE.lock().unwrap().active_sheet = sheet_name;
}

/// Get active sheet
pub fn get_active_sheet() -> String {
    ENGINE_STATE.lock().unwrap().active_sheet.clone()
}

/// Convert Excel address like "A1" to (row, col) indices
pub fn address_to_indices(address: &str) -> Result<(i32, i32), String> {
    let address = address.trim().to_uppercase();
    
    // Split into column letters and row number
    let col_end = address.chars().take_while(|c| c.is_alphabetic()).count();
    if col_end == 0 {
        return Err(format!("Invalid address: {}", address));
    }
    
    let col_str = &address[..col_end];
    let row_str = &address[col_end..];
    
    // Parse row (1-based in Excel, 0-based in engine)
    let row: i32 = row_str.parse()
        .map_err(|_| format!("Invalid row number: {}", row_str))?;
    let row_idx = row - 1;
    
    // Parse column (A=0, B=1, ..., Z=25, AA=26, etc.)
    let mut col_idx: i32 = 0;
    for ch in col_str.chars() {
        col_idx = col_idx * 26 + (ch as i32 - 'A' as i32 + 1);
    }
    col_idx -= 1; // Convert to 0-based
    
    Ok((row_idx, col_idx))
}

/// Set cell value in the native engine
/// Supports both single cell (A1) and range (A1:A12) addresses
pub fn set_cell_value(address: &str, value: &str) -> Result<(), String> {
    let state = ENGINE_STATE.lock().unwrap();
    
    if !state.initialized {
        // Engine not initialized - stub mode
        // Just log, don't store anything - cells are engine's responsibility
        // eprintln!("⚠️  [Stub] Set {}!{} = '{}' (engine not initialized)", state.active_sheet, address, value);
        return Ok(());
    }
    
    let workbook_id = state.workbook_id.as_ref()
        .ok_or("No workbook ID")?;
    let sheet_name = &state.active_sheet;
    
    // Check if this is a range (contains :)
    if address.contains(':') {
        // It's a range - expand and set all cells
        let parts: Vec<&str> = address.split(':').collect();
        if parts.len() != 2 {
            return Err(format!("Invalid range: {}", address));
        }
        
        let (start_row, start_col) = address_to_indices(parts[0])?;
        let (end_row, end_col) = address_to_indices(parts[1])?;
        
        // Make sure we handle ranges in any direction
        let min_row = start_row.min(end_row);
        let max_row = start_row.max(end_row);
        let min_col = start_col.min(end_col);
        let max_col = start_col.max(end_col);
        
        for row in min_row..=max_row {
            for col in min_col..=max_col {
                unsafe {
                    let wb_cstr = CString::new(workbook_id.as_str()).map_err(|e| e.to_string())?;
                    let sheet_cstr = CString::new(sheet_name.as_str()).map_err(|e| e.to_string())?;
                    let value_cstr = CString::new(value).map_err(|e| e.to_string())?;
                    
                    let success = EngineInterface_SetCellValue(
                        wb_cstr.as_ptr(),
                        sheet_cstr.as_ptr(),
                        row,
                        col,
                        value_cstr.as_ptr()
                    );
                    
                    if !success {
                        return Err(format!("Failed to set cell ({}, {})", row, col));
                    }
                }
            }
        }
        return Ok(());
    }
    
    // Single cell
    let (row, col) = address_to_indices(address)?;
    
    unsafe {
        let wb_cstr = CString::new(workbook_id.as_str()).map_err(|e| e.to_string())?;
        let sheet_cstr = CString::new(sheet_name.as_str()).map_err(|e| e.to_string())?;
        let value_cstr = CString::new(value).map_err(|e| e.to_string())?;
        
        let success = EngineInterface_SetCellValue(
            wb_cstr.as_ptr(),
            sheet_cstr.as_ptr(),
            row,
            col,
            value_cstr.as_ptr()
        );
        
        if success {
            // eprintln!("✅ Set {}!{} = '{}'", sheet_name, address, value);
            Ok(())
        } else {
            Err(format!("Failed to set cell {}", address))
        }
    }
}

/// Get cell value from the native engine
pub fn get_cell_value(address: &str) -> Result<String, String> {
    let state = ENGINE_STATE.lock().unwrap();
    let sheet_name = &state.active_sheet;
    
    if !state.initialized {
        // Engine not initialized - stub mode 
        // Return empty string - cells are engine's responsibility
        eprintln!("⚠️  [Stub] Get {}!{} = '' (engine not initialized)", sheet_name, address);
        return Ok(String::new());
    }
    
    let workbook_id = state.workbook_id.as_ref()
        .ok_or("No workbook ID")?;
    
    let (row, col) = address_to_indices(address)?;
    
    unsafe {
        let wb_cstr = CString::new(workbook_id.as_str()).map_err(|e| e.to_string())?;
        let sheet_cstr = CString::new(sheet_name.as_str()).map_err(|e| e.to_string())?;
        
        let value_ptr = EngineInterface_GetCellValue(
            wb_cstr.as_ptr(),
            sheet_cstr.as_ptr(),
            row,
            col
        );
        
        if value_ptr.is_null() {
            return Ok(String::new()); // Empty cell
        }
        
        let value = CStr::from_ptr(value_ptr)
            .to_str()
            .map_err(|e| e.to_string())?
            .to_owned();
        
        free(value_ptr as *mut libc::c_void);
        
        // eprintln!("📖 Get {}!{} = '{}'", sheet_name, address, value);
        Ok(value)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_address_parsing() {
        assert_eq!(address_to_indices("A1").unwrap(), (0, 0));
        assert_eq!(address_to_indices("B1").unwrap(), (0, 1));
        assert_eq!(address_to_indices("Z1").unwrap(), (0, 25));
        assert_eq!(address_to_indices("AA1").unwrap(), (0, 26));
        assert_eq!(address_to_indices("A10").unwrap(), (9, 0));
        assert_eq!(address_to_indices("C5").unwrap(), (4, 2));
    }
    
}