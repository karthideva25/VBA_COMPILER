// Tests for VBA Date/Time Functions
//
// This test file covers VBA date/time functions including:
// - Year, Month, Day, Hour, Minute, Second
// - Weekday, WeekdayName, MonthName
// - DateSerial, TimeSerial, DateValue, TimeValue
// - DateAdd, DateDiff, DatePart
// - FormatDateTime, IsDate
//
// Note: Now(), Date(), Time(), Timer() are not tested as they return current time

use tree_sitter::Parser;
use vba_parser::language as tree_sitter_vba;
use vba_utils::Context;
use vba_utils::vm::ProgramExecutor;
use vba_utils::ast::build_ast;

/// Helper to run VBA code and capture output
fn run_vba(code: &str) -> Vec<String> {
    let mut parser = Parser::new();
    parser.set_language(tree_sitter_vba()).expect("Failed to set VBA language");
    let tree = parser.parse(code, None).expect("Failed to parse VBA code");
    let root_node = tree.root_node();
    let program = build_ast(root_node, code);
    
    let mut ctx = Context::new();
    let executor = ProgramExecutor::new(program);
    let _ = executor.execute(&mut ctx);
    ctx.output.clone()
}

/// Helper to run VBA code and get first output value
fn run_vba_first(code: &str) -> String {
    let output = run_vba(code);
    output.first().cloned().unwrap_or_default()
}

// ============================================================
// DATESERIAL TESTS
// ============================================================

#[test]
fn test_dateserial_basic() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 15)
            MsgBox Year(d) & "-" & Month(d) & "-" & Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2026-1-15");
}

#[test]
fn test_dateserial_month_overflow() {
    // Month 13 = Jan of next year
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 13, 1)
            MsgBox Year(d) & "-" & Month(d) & "-" & Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2027-1-1");
}

#[test]
fn test_dateserial_month_negative() {
    // Month 0 = Dec of previous year
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 0, 15)
            MsgBox Year(d) & "-" & Month(d) & "-" & Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2025-12-15");
}

#[test]
fn test_dateserial_day_overflow() {
    // Day 32 of Jan = Feb 1
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 32)
            MsgBox Month(d) & "-" & Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2-1");
}

// ============================================================
// YEAR / MONTH / DAY TESTS
// ============================================================

#[test]
fn test_year_basic() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 6, 15)
            MsgBox Year(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2026");
}

#[test]
fn test_month_basic() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 6, 15)
            MsgBox Month(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "6");
}

#[test]
fn test_day_basic() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 6, 15)
            MsgBox Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "15");
}

// ============================================================
// TIMESERIAL TESTS
// ============================================================

#[test]
fn test_timeserial_basic() {
    let code = r#"
        Sub AutoOpen()
            Dim t As Variant
            t = TimeSerial(14, 30, 45)
            MsgBox Hour(t) & ":" & Minute(t) & ":" & Second(t)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "14:30:45");
}

#[test]
fn test_timeserial_minute_overflow() {
    // 90 minutes = 1 hour 30 minutes
    let code = r#"
        Sub AutoOpen()
            Dim t As Variant
            t = TimeSerial(0, 90, 0)
            MsgBox Hour(t) & ":" & Minute(t)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1:30");
}

#[test]
fn test_timeserial_negative_hour() {
    // -1 hour from midnight wraps to previous day, so with +30 min = 23:30
    let code = r#"
        Sub AutoOpen()
            Dim t As Variant
            t = TimeSerial(-1, 30, 0)
            MsgBox Hour(t) & ":" & Minute(t)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "23:30");
}

#[test]
fn test_timeserial_second_overflow() {
    // 90 seconds = 1 min 30 sec
    let code = r#"
        Sub AutoOpen()
            Dim t As Variant
            t = TimeSerial(0, 0, 90)
            MsgBox Minute(t) & ":" & Second(t)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1:30");
}

// ============================================================
// HOUR / MINUTE / SECOND TESTS
// ============================================================

#[test]
fn test_hour_basic() {
    let code = r#"
        Sub AutoOpen()
            Dim t As Variant
            t = TimeSerial(14, 30, 45)
            MsgBox Hour(t)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "14");
}

#[test]
fn test_minute_basic() {
    let code = r#"
        Sub AutoOpen()
            Dim t As Variant
            t = TimeSerial(14, 30, 45)
            MsgBox Minute(t)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "30");
}

#[test]
fn test_second_basic() {
    let code = r#"
        Sub AutoOpen()
            Dim t As Variant
            t = TimeSerial(14, 30, 45)
            MsgBox Second(t)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "45");
}

// ============================================================
// WEEKDAY TESTS
// ============================================================

#[test]
fn test_weekday_sunday() {
    // Jan 4, 2026 is a Sunday
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 4)
            MsgBox Weekday(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1");
}

#[test]
fn test_weekday_saturday() {
    // Jan 3, 2026 is a Saturday
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 3)
            MsgBox Weekday(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "7");
}

#[test]
fn test_weekday_monday() {
    // Jan 5, 2026 is a Monday
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 5)
            MsgBox Weekday(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2");
}

// ============================================================
// WEEKDAYNAME / MONTHNAME TESTS
// ============================================================

#[test]
fn test_weekdayname_full() {
    let code = r#"
        Sub AutoOpen()
            MsgBox WeekdayName(1)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Sunday");
}

#[test]
fn test_weekdayname_abbreviated() {
    let code = r#"
        Sub AutoOpen()
            MsgBox WeekdayName(2, True)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Mon");
}

#[test]
fn test_monthname_full() {
    let code = r#"
        Sub AutoOpen()
            MsgBox MonthName(1)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "January");
}

#[test]
fn test_monthname_abbreviated() {
    let code = r#"
        Sub AutoOpen()
            MsgBox MonthName(12, True)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "Dec");
}

#[test]
fn test_monthname_all_months() {
    let code = r#"
        Sub AutoOpen()
            MsgBox MonthName(7)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "July");
}

// ============================================================
// DATEADD TESTS
// ============================================================

#[test]
fn test_dateadd_days() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 1)
            d = DateAdd("d", 10, d)
            MsgBox Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "11");
}

#[test]
fn test_dateadd_months() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 15)
            d = DateAdd("m", 2, d)
            MsgBox Month(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "3");
}

#[test]
fn test_dateadd_years() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 15)
            d = DateAdd("yyyy", 5, d)
            MsgBox Year(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2031");
}

#[test]
fn test_dateadd_weeks() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 1)
            d = DateAdd("ww", 2, d)
            MsgBox Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "15");
}

#[test]
fn test_dateadd_hours() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Variant
            d = DateSerial(2026, 1, 1)
            d = DateAdd("h", 25, d)
            MsgBox Day(d) & " " & Hour(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2 1");
}

#[test]
fn test_dateadd_minutes() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Variant
            d = DateSerial(2026, 1, 1)
            d = DateAdd("n", 90, d)
            MsgBox Hour(d) & ":" & Minute(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1:30");
}

// ============================================================
// DATEDIFF TESTS
// ============================================================

#[test]
fn test_datediff_days() {
    let code = r#"
        Sub AutoOpen()
            Dim d1 As Date, d2 As Date
            d1 = DateSerial(2026, 1, 1)
            d2 = DateSerial(2026, 1, 15)
            MsgBox DateDiff("d", d1, d2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "14");
}

#[test]
fn test_datediff_months() {
    let code = r#"
        Sub AutoOpen()
            Dim d1 As Date, d2 As Date
            d1 = DateSerial(2026, 1, 1)
            d2 = DateSerial(2026, 6, 1)
            MsgBox DateDiff("m", d1, d2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "5");
}

#[test]
fn test_datediff_years() {
    let code = r#"
        Sub AutoOpen()
            Dim d1 As Date, d2 As Date
            d1 = DateSerial(2020, 1, 1)
            d2 = DateSerial(2026, 1, 1)
            MsgBox DateDiff("yyyy", d1, d2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "6");
}

#[test]
fn test_datediff_weeks() {
    let code = r#"
        Sub AutoOpen()
            Dim d1 As Date, d2 As Date
            d1 = DateSerial(2026, 1, 1)
            d2 = DateSerial(2026, 1, 22)
            MsgBox DateDiff("ww", d1, d2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "3");
}

#[test]
fn test_datediff_negative() {
    // d1 > d2 should give negative result
    let code = r#"
        Sub AutoOpen()
            Dim d1 As Date, d2 As Date
            d1 = DateSerial(2026, 1, 15)
            d2 = DateSerial(2026, 1, 1)
            MsgBox DateDiff("d", d1, d2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "-14");
}

// ============================================================
// DATEPART TESTS
// ============================================================

#[test]
fn test_datepart_year() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 6, 15)
            MsgBox DatePart("yyyy", d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2026");
}

#[test]
fn test_datepart_month() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 6, 15)
            MsgBox DatePart("m", d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "6");
}

#[test]
fn test_datepart_day() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 6, 15)
            MsgBox DatePart("d", d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "15");
}

#[test]
fn test_datepart_quarter() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 7, 15)
            MsgBox DatePart("q", d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "3");
}

#[test]
fn test_datepart_weekday() {
    // Jan 4, 2026 is Sunday = 1
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 4)
            MsgBox DatePart("w", d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "1");
}

#[test]
fn test_datepart_day_of_year() {
    // Feb 1, 2026 = day 32 of year
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 2, 1)
            MsgBox DatePart("y", d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "32");
}

// ============================================================
// DATEVALUE / TIMEVALUE TESTS
// ============================================================

#[test]
fn test_datevalue_basic() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateValue("2026-01-15")
            MsgBox Year(d) & "-" & Month(d) & "-" & Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2026-1-15");
}

#[test]
fn test_timevalue_basic() {
    let code = r#"
        Sub AutoOpen()
            Dim t As Variant
            t = TimeValue("14:30:45")
            MsgBox Hour(t) & ":" & Minute(t) & ":" & Second(t)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "14:30:45");
}

// ============================================================
// ISDATE TESTS
// ============================================================

#[test]
fn test_isdate_true() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 15)
            MsgBox IsDate(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "True");
}

#[test]
fn test_isdate_string_valid() {
    let code = r#"
        Sub AutoOpen()
            MsgBox IsDate("2026-01-15")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "True");
}

#[test]
fn test_isdate_string_invalid() {
    let code = r#"
        Sub AutoOpen()
            MsgBox IsDate("not a date")
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "False");
}

#[test]
fn test_isdate_number() {
    let code = r#"
        Sub AutoOpen()
            MsgBox IsDate(12345)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "False");
}

// ============================================================
// FORMATDATETIME TESTS
// ============================================================

#[test]
fn test_formatdatetime_general() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 15)
            MsgBox FormatDateTime(d, 0)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "01/15/2026");
}

#[test]
fn test_formatdatetime_long() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 15)
            MsgBox FormatDateTime(d, 1)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "January 15, 2026");
}

#[test]
fn test_formatdatetime_short() {
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 15)
            MsgBox FormatDateTime(d, 2)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "01/15/2026");
}

// ============================================================
// COMBINED DATE OPERATIONS
// ============================================================

#[test]
fn test_combined_date_arithmetic() {
    // Add 1 month and 5 days
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 1, 10)
            d = DateAdd("m", 1, d)
            d = DateAdd("d", 5, d)
            MsgBox Month(d) & "-" & Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2-15");
}

#[test]
fn test_calculate_age_in_days() {
    let code = r#"
        Sub AutoOpen()
            Dim birthDate As Date, currentDate As Date
            birthDate = DateSerial(2000, 1, 1)
            currentDate = DateSerial(2026, 1, 1)
            MsgBox DateDiff("d", birthDate, currentDate)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "9497");
}

#[test]
fn test_end_of_month() {
    // Getting last day of Feb 2026 (non-leap year)
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2026, 3, 0)
            MsgBox Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "28");
}

#[test]
fn test_leap_year_feb() {
    // 2024 is a leap year
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2024, 3, 0)
            MsgBox Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "29");
}

// ============================================================
// EDGE CASES
// ============================================================

#[test]
fn test_midnight_time() {
    let code = r#"
        Sub AutoOpen()
            Dim t As Variant
            t = TimeSerial(0, 0, 0)
            MsgBox Hour(t) & ":" & Minute(t) & ":" & Second(t)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "0:0:0");
}

#[test]
fn test_end_of_day_time() {
    let code = r#"
        Sub AutoOpen()
            Dim t As Variant
            t = TimeSerial(23, 59, 59)
            MsgBox Hour(t) & ":" & Minute(t) & ":" & Second(t)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "23:59:59");
}

#[test]
fn test_year_boundary() {
    // Dec 31 + 1 day = Jan 1 next year
    let code = r#"
        Sub AutoOpen()
            Dim d As Date
            d = DateSerial(2025, 12, 31)
            d = DateAdd("d", 1, d)
            MsgBox Year(d) & "-" & Month(d) & "-" & Day(d)
        End Sub
    "#;
    assert_eq!(run_vba_first(code), "2026-1-1");
}
