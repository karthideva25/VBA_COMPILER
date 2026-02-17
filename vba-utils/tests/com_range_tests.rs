// COMtests/range_tests.rs
// ============================================================================
// Tests for Excel Range COM Object
// ============================================================================

use vba_utils::host::excel::objects::range::{ExcelRange, RangeBuilder, column_index_to_letter, indices_to_address};

#[test]
fn test_range_creation() {
    let rng = ExcelRange::new("A1");
    assert_eq!(rng.address, "A1");
    assert!(rng.sheet_name.is_none());
}

#[test]
fn test_range_with_sheet() {
    let rng = ExcelRange::new("Sheet1!A1:B5");
    assert_eq!(rng.address, "A1:B5");
    assert_eq!(rng.sheet_name, Some("Sheet1".to_string()));
}

#[test]
fn test_range_full_address() {
    let rng = ExcelRange::with_sheet("A1:B5", "My Sheet");
    assert_eq!(rng.full_address(), "'My Sheet'!A1:B5");
}

#[test]
fn test_is_multi_cell() {
    assert!(!ExcelRange::new("A1").is_multi_cell());
    assert!(ExcelRange::new("A1:B5").is_multi_cell());
}

#[test]
fn test_column_letter() {
    assert_eq!(column_index_to_letter(0), "A");
    assert_eq!(column_index_to_letter(25), "Z");
    assert_eq!(column_index_to_letter(26), "AA");
    assert_eq!(column_index_to_letter(27), "AB");
    assert_eq!(column_index_to_letter(701), "ZZ");
}

#[test]
fn test_indices_to_address() {
    assert_eq!(indices_to_address(0, 0), "A1");
    assert_eq!(indices_to_address(4, 2), "C5");
    assert_eq!(indices_to_address(99, 26), "AA100");
}

#[test]
fn test_range_builder() {
    let rng = RangeBuilder::from_cell(0, 0)
        .to_cell(4, 2)
        .on_sheet("Data")
        .build();
    assert_eq!(rng.full_address(), "Data!A1:C5");
}

#[test]
fn test_range_bounds() {
    let rng = ExcelRange::new("B2:D5");
    let bounds = rng.get_bounds().unwrap();
    assert_eq!(bounds, ((1, 1), (4, 3))); // 0-based: B=1, 2=row1, D=3, 5=row4
}

#[test]
fn test_range_row_count() {
    let rng = ExcelRange::new("A1:A10");
    assert_eq!(rng.row_count().unwrap(), 10);
    
    let single = ExcelRange::new("A1");
    assert_eq!(single.row_count().unwrap(), 1);
}

#[test]
fn test_range_col_count() {
    let rng = ExcelRange::new("A1:E1");
    assert_eq!(rng.col_count().unwrap(), 5);
    
    let single = ExcelRange::new("A1");
    assert_eq!(single.col_count().unwrap(), 1);
}

#[test]
fn test_range_cell_count() {
    let rng = ExcelRange::new("A1:C5");
    assert_eq!(rng.cell_count().unwrap(), 15); // 3 cols * 5 rows
}

#[test]
fn test_range_offset() {
    let rng = ExcelRange::new("A1");
    let offset_rng = rng.offset(2, 3).unwrap();
    assert_eq!(offset_rng.address, "D3");
    
    // Multi-cell range offset
    let multi = ExcelRange::new("A1:B2");
    let offset_multi = multi.offset(1, 1).unwrap();
    assert_eq!(offset_multi.address, "B2:C3");
}

#[test]
fn test_range_resize() {
    let rng = ExcelRange::new("A1");
    let resized = rng.resize(Some(3), Some(2)).unwrap();
    assert_eq!(resized.address, "A1:B3");
    
    // Resize to single cell
    let multi = ExcelRange::new("A1:D5");
    let single = multi.resize(Some(1), Some(1)).unwrap();
    assert_eq!(single.address, "A1");
}

#[test]
fn test_range_cells() {
    let rng = ExcelRange::new("B2:D5");
    
    // Cell(1,1) should be the top-left = B2
    let cell = rng.cells(1, 1).unwrap();
    assert_eq!(cell.address, "B2");
    
    // Cell(2,2) should be C3
    let cell2 = rng.cells(2, 2).unwrap();
    assert_eq!(cell2.address, "C3");
}

#[test]
fn test_range_top_left_bottom_right() {
    let rng = ExcelRange::new("B2:D5");
    assert_eq!(rng.top_left().unwrap(), "B2");
    assert_eq!(rng.bottom_right().unwrap(), "D5");
}
