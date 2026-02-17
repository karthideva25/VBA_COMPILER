//! VBA Date/Time Functions
//! 
//! This module contains all VBA date and time functions including:
//! - Now, Date, Time
//! - Year, Month, Day, Hour, Minute, Second
//! - Weekday, WeekdayName, MonthName
//! - DateSerial, TimeSerial, DateValue, TimeValue
//! - DateAdd, DateDiff, DatePart
//! - FormatDateTime
//! - IsDate
//!
//! All date/time functions respect the user's timezone configured in
//! `Context.runtime_config.timezone`. This is set at session start by
//! the application layer.

use anyhow::Result;
use chrono::{Datelike, NaiveDate, Timelike, Utc};
use chrono::TimeZone as _;  // Extension trait for with_timezone()
use crate::ast::Expression;
use crate::context::{Context, Value};
use crate::interpreter::evaluate_expression;

/// Handle date/time-related builtin function calls
pub(crate) fn handle_datetime_function(function: &str, args: &[Expression], ctx: &mut Context) -> Result<Option<Value>> {
    match function {
        // ============================================================
        // CURRENT DATE/TIME (uses user's timezone from RuntimeConfig)
        // ============================================================

        // NOW — Returns current date and time in user's timezone
        "now" => {
            let tz = ctx.runtime_config.timezone;
            let now_in_tz = Utc::now().with_timezone(&tz);
            let datetime = now_in_tz.naive_local();
            Ok(Some(Value::DateTime(datetime)))
        }

        // DATE — Returns the current system date in user's timezone
        "date" => {
            let tz = ctx.runtime_config.timezone;
            let now_in_tz = Utc::now().with_timezone(&tz);
            let today = now_in_tz.date_naive();
            Ok(Some(Value::Date(today)))
        }

        // TIME — Returns the current system time in user's timezone
        "time" => {
            let tz = ctx.runtime_config.timezone;
            let now_in_tz = Utc::now().with_timezone(&tz);
            let time = now_in_tz.time();
            Ok(Some(Value::Time(time)))
        }

        // TIMER — Returns seconds since midnight as a Single (float)
        "timer" => {
            let tz = ctx.runtime_config.timezone;
            let now_in_tz = Utc::now().with_timezone(&tz);
            let time = now_in_tz.time();
            // Calculate seconds since midnight including fractional seconds
            let seconds = time.hour() as f64 * 3600.0 
                        + time.minute() as f64 * 60.0 
                        + time.second() as f64
                        + (time.nanosecond() as f64 / 1_000_000_000.0);
            Ok(Some(Value::Single(seconds as f32)))
        }

        // ============================================================
        // DATE PART EXTRACTION
        // ============================================================

        // YEAR — Returns the year from a date
        "year" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::Date(d) => Ok(Some(Value::Integer(d.year() as i64))),
                Value::DateTime(dt) => Ok(Some(Value::Integer(dt.date().year() as i64))),
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // MONTH — Returns the month from a date (1-12)
        "month" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::Date(d) => Ok(Some(Value::Integer(d.month() as i64))),
                Value::DateTime(dt) => Ok(Some(Value::Integer(dt.date().month() as i64))),
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // DAY — Returns the day of month from a date (1-31)
        "day" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::Date(d) => Ok(Some(Value::Integer(d.day() as i64))),
                Value::DateTime(dt) => Ok(Some(Value::Integer(dt.date().day() as i64))),
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // WEEKDAY — Returns the day of week (1=Sunday, 7=Saturday)
        "weekday" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::Date(d) => {
                    // chrono: Mon=0, Sun=6
                    // VBA default: Sun=1, Sat=7
                    let weekday = d.weekday().num_days_from_sunday() + 1;
                    Ok(Some(Value::Integer(weekday as i64)))
                }
                Value::DateTime(dt) => {
                    let weekday = dt.date().weekday().num_days_from_sunday() + 1;
                    Ok(Some(Value::Integer(weekday as i64)))
                }
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // HOUR — Returns hour (0-23)
        "hour" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::DateTime(dt) => Ok(Some(Value::Integer(dt.hour() as i64))),
                Value::Time(t) => Ok(Some(Value::Integer(t.hour() as i64))),
                Value::Date(_) => Ok(Some(Value::Integer(0))), // Date has no time component
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // MINUTE — Returns minute (0-59)
        "minute" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::DateTime(dt) => Ok(Some(Value::Integer(dt.minute() as i64))),
                Value::Time(t) => Ok(Some(Value::Integer(t.minute() as i64))),
                Value::Date(_) => Ok(Some(Value::Integer(0))),
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // SECOND — Returns second (0-59)
        "second" => {
            if args.is_empty() {
                return Ok(Some(Value::Integer(0)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            match val {
                Value::DateTime(dt) => Ok(Some(Value::Integer(dt.second() as i64))),
                Value::Time(t) => Ok(Some(Value::Integer(t.second() as i64))),
                Value::Date(_) => Ok(Some(Value::Integer(0))),
                _ => Ok(Some(Value::Integer(0)))
            }
        }

        // ============================================================
        // DATE CREATION
        // ============================================================

        // DATESERIAL — DateSerial(year, month, day)
        // VBA allows overflow: DateSerial(2026, 13, 1) = Jan 1, 2027
        "dateserial" => {
            if args.len() != 3 {
                return Ok(Some(Value::Empty));
            }
            let year_val = evaluate_expression(&args[0], ctx)?;
            let month_val = evaluate_expression(&args[1], ctx)?;
            let day_val = evaluate_expression(&args[2], ctx)?;
            
            let year = match &year_val { 
                Value::Integer(i) => *i as i32, 
                Value::Long(i) => *i,
                Value::Double(d) => *d as i32,
                _ => return Ok(Some(Value::Empty)) 
            };
            let month = match &month_val { 
                Value::Integer(i) => *i as i32, 
                Value::Long(i) => *i as i32,
                Value::Double(d) => *d as i32,
                _ => return Ok(Some(Value::Empty)) 
            };
            let day = match &day_val { 
                Value::Integer(i) => *i as i32, 
                Value::Long(i) => *i as i32,
                Value::Double(d) => *d as i32,
                _ => return Ok(Some(Value::Empty)) 
            };
            
            // VBA behavior: months and days can overflow/underflow
            // Start with Jan 1 of the year, then add months-1, then add days-1
            if let Some(base_date) = NaiveDate::from_ymd_opt(year, 1, 1) {
                // Add months (month - 1 because we started at January)
                let months_to_add = month - 1;  // e.g., month=3 means add 2 months from Jan
                let with_months = if months_to_add >= 0 {
                    base_date.checked_add_months(chrono::Months::new(months_to_add as u32))
                } else {
                    // Negative months: go back from Jan 1
                    // month=0 means Dec of prev year, month=-1 means Nov of prev year
                    base_date.checked_sub_months(chrono::Months::new((-months_to_add) as u32))
                };
                
                // Add days (day - 1 because we started at day 1)
                if let Some(date_with_months) = with_months {
                    let with_days = date_with_months.checked_add_signed(chrono::Duration::days((day - 1) as i64));
                    if let Some(final_date) = with_days {
                        return Ok(Some(Value::Date(final_date)));
                    }
                }
            }
            
            // Fallback: try direct construction
            if month >= 1 && month <= 12 && day >= 1 && day <= 31 {
                if let Some(date) = NaiveDate::from_ymd_opt(year, month as u32, day as u32) {
                    return Ok(Some(Value::Date(date)));
                }
            }
            
            Ok(Some(Value::Empty))
        }

        // DATEVALUE — DateValue(date_string)
        "datevalue" => {
            if args.is_empty() {
                return Ok(Some(Value::Empty));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let date_str = match val {
                Value::String(s) => s,
                Value::Date(d) => return Ok(Some(Value::Date(d))),
                _ => return Ok(Some(Value::Empty))
            };
            
            // Try common date formats
            let formats = [
                "%Y-%m-%d",
                "%m/%d/%Y",
                "%d/%m/%Y",
                "%Y/%m/%d",
                "%B %d, %Y",
                "%b %d, %Y",
            ];
            
            for fmt in formats.iter() {
                if let Ok(date) = NaiveDate::parse_from_str(&date_str, fmt) {
                    return Ok(Some(Value::Date(date)));
                }
            }
            
            Ok(Some(Value::Empty))
        }

        // TIMEVALUE — TimeValue(time_string) - Converts a string to a Time value
        "timevalue" => {
            if args.is_empty() {
                return Ok(Some(Value::Empty));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let time_str = match val {
                Value::String(s) => s,
                Value::Time(t) => return Ok(Some(Value::Time(t))),
                _ => return Ok(Some(Value::Empty))
            };
            
            // Try common time formats
            let formats = [
                "%H:%M:%S",      // 14:30:00
                "%H:%M",        // 14:30
                "%I:%M:%S %p",  // 02:30:00 PM
                "%I:%M %p",     // 02:30 PM
            ];
            
            for fmt in formats.iter() {
                if let Ok(time) = chrono::NaiveTime::parse_from_str(&time_str, fmt) {
                    return Ok(Some(Value::Time(time)));
                }
            }
            
            Ok(Some(Value::Empty))
        }

        // TIMESERIAL — TimeSerial(hour, minute, second) - Creates a Time value
        // VBA allows overflow: TimeSerial(0, 90, 0) = 01:30:00
        // VBA allows negative: TimeSerial(-1, 30, 0) = 22:30:00
        "timeserial" => {
            if args.len() != 3 {
                return Ok(Some(Value::Empty));
            }
            let hour_val = evaluate_expression(&args[0], ctx)?;
            let min_val = evaluate_expression(&args[1], ctx)?;
            let sec_val = evaluate_expression(&args[2], ctx)?;
            
            let hour = match &hour_val { 
                Value::Integer(i) => *i, 
                Value::Long(i) => *i as i64,
                Value::Double(d) => *d as i64,
                _ => return Ok(Some(Value::Empty)) 
            };
            let minute = match &min_val { 
                Value::Integer(i) => *i, 
                Value::Long(i) => *i as i64,
                Value::Double(d) => *d as i64,
                _ => return Ok(Some(Value::Empty)) 
            };
            let second = match &sec_val { 
                Value::Integer(i) => *i, 
                Value::Long(i) => *i as i64,
                Value::Double(d) => *d as i64,
                _ => return Ok(Some(Value::Empty)) 
            };
            
            // Convert everything to total seconds, then normalize to 0..86400
            let total_seconds = hour * 3600 + minute * 60 + second;
            // VBA wraps around at day boundaries (86400 seconds = 24 hours)
            let normalized = ((total_seconds % 86400) + 86400) % 86400;  // Handle negatives
            
            let final_hour = (normalized / 3600) as u32;
            let final_minute = ((normalized % 3600) / 60) as u32;
            let final_second = (normalized % 60) as u32;
            
            if let Some(time) = chrono::NaiveTime::from_hms_opt(final_hour, final_minute, final_second) {
                Ok(Some(Value::Time(time)))
            } else {
                Ok(Some(Value::Empty))
            }
        }

        // ============================================================
        // DATE ARITHMETIC
        // ============================================================

        // DATEADD — DateAdd(interval, number, date)
        "dateadd" => {
            if args.len() != 3 {
                return Ok(Some(Value::Empty));
            }
            let interval_val = evaluate_expression(&args[0], ctx)?;
            let number_val = evaluate_expression(&args[1], ctx)?;
            let date_val = evaluate_expression(&args[2], ctx)?;
            
            let interval = match interval_val { Value::String(s) => s.to_lowercase(), _ => return Ok(Some(Value::Empty)) };
            let number = match &number_val { 
                Value::Integer(i) => *i, 
                Value::Long(i) => *i as i64,
                Value::Double(d) => *d as i64,
                Value::Single(s) => *s as i64,
                _ => return Ok(Some(Value::Empty)) 
            };
            
            // Extract both date and optional time component
            let (date, time_opt) = match date_val { 
                Value::Date(d) => (d, None), 
                Value::DateTime(dt) => (dt.date(), Some(dt.time())),
                _ => return Ok(Some(Value::Empty)) 
            };
            
            // For time intervals (h, n, s), we need to work with DateTime
            match interval.as_str() {
                "h" => {
                    // Add hours - result is DateTime
                    let base_time = time_opt.unwrap_or_else(|| chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap());
                    let dt = chrono::NaiveDateTime::new(date, base_time);
                    if let Some(new_dt) = dt.checked_add_signed(chrono::Duration::hours(number)) {
                        return Ok(Some(Value::DateTime(new_dt)));
                    }
                    return Ok(Some(Value::Empty));
                }
                "n" => {
                    // Add minutes - result is DateTime
                    let base_time = time_opt.unwrap_or_else(|| chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap());
                    let dt = chrono::NaiveDateTime::new(date, base_time);
                    if let Some(new_dt) = dt.checked_add_signed(chrono::Duration::minutes(number)) {
                        return Ok(Some(Value::DateTime(new_dt)));
                    }
                    return Ok(Some(Value::Empty));
                }
                "s" => {
                    // Add seconds - result is DateTime
                    let base_time = time_opt.unwrap_or_else(|| chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap());
                    let dt = chrono::NaiveDateTime::new(date, base_time);
                    if let Some(new_dt) = dt.checked_add_signed(chrono::Duration::seconds(number)) {
                        return Ok(Some(Value::DateTime(new_dt)));
                    }
                    return Ok(Some(Value::Empty));
                }
                _ => {}
            }
            
            let result = match interval.as_str() {
                "yyyy" | "y" => {
                    // Add years
                    NaiveDate::from_ymd_opt(date.year() + number as i32, date.month(), date.day())
                }
                "m" => {
                    // Add months
                    let total_months = date.year() * 12 + date.month() as i32 + number as i32;
                    let new_year = (total_months - 1) / 12;
                    let new_month = ((total_months - 1) % 12 + 1) as u32;
                    NaiveDate::from_ymd_opt(new_year, new_month, date.day().min(28))
                }
                "d" => {
                    // Add days
                    date.checked_add_signed(chrono::Duration::days(number))
                }
                "ww" | "w" => {
                    // Add weeks
                    date.checked_add_signed(chrono::Duration::weeks(number))
                }
                _ => None
            };
            
            match result {
                Some(d) => Ok(Some(Value::Date(d))),
                None => Ok(Some(Value::Empty))
            }
        }

        // DATEDIFF — DateDiff(interval, date1, date2)
        "datediff" => {
            if args.len() != 3 {
                return Ok(Some(Value::Integer(0)));
            }
            let interval_val = evaluate_expression(&args[0], ctx)?;
            let date1_val = evaluate_expression(&args[1], ctx)?;
            let date2_val = evaluate_expression(&args[2], ctx)?;
            
            let interval = match interval_val { Value::String(s) => s.to_lowercase(), _ => return Ok(Some(Value::Integer(0))) };
            
            // Extract full DateTime for time-based intervals
            let dt1 = match &date1_val { 
                Value::Date(d) => chrono::NaiveDateTime::new(*d, chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap()), 
                Value::DateTime(dt) => *dt,
                _ => return Ok(Some(Value::Integer(0))) 
            };
            let dt2 = match &date2_val { 
                Value::Date(d) => chrono::NaiveDateTime::new(*d, chrono::NaiveTime::from_hms_opt(0, 0, 0).unwrap()), 
                Value::DateTime(dt) => *dt,
                _ => return Ok(Some(Value::Integer(0))) 
            };
            
            let date1 = dt1.date();
            let date2 = dt2.date();
            
            let diff = match interval.as_str() {
                "yyyy" | "y" => (date2.year() - date1.year()) as i64,
                "m" => {
                    let months1 = date1.year() * 12 + date1.month() as i32;
                    let months2 = date2.year() * 12 + date2.month() as i32;
                    (months2 - months1) as i64
                }
                "d" => (date2 - date1).num_days(),
                "ww" | "w" => (date2 - date1).num_weeks(),
                "h" => {
                    // Difference in hours
                    let duration = dt2 - dt1;
                    duration.num_hours()
                }
                "n" => {
                    // Difference in minutes
                    let duration = dt2 - dt1;
                    duration.num_minutes()
                }
                "s" => {
                    // Difference in seconds
                    let duration = dt2 - dt1;
                    duration.num_seconds()
                }
                _ => 0
            };
            
            Ok(Some(Value::Integer(diff)))
        }

        // DATEPART — DatePart(interval, date)
        "datepart" => {
            if args.len() < 2 {
                return Ok(Some(Value::Integer(0)));
            }
            let interval_val = evaluate_expression(&args[0], ctx)?;
            let date_val = evaluate_expression(&args[1], ctx)?;
            
            let interval = match interval_val { Value::String(s) => s.to_lowercase(), _ => return Ok(Some(Value::Integer(0))) };
            let date = match date_val { 
                Value::Date(d) => d, 
                Value::DateTime(dt) => dt.date(),
                _ => return Ok(Some(Value::Integer(0))) 
            };
            
            let part = match interval.as_str() {
                "yyyy" => date.year() as i64,
                "q" => ((date.month() - 1) / 3 + 1) as i64,  // Quarter
                "m" => date.month() as i64,
                "y" => date.ordinal() as i64,  // Day of year
                "d" => date.day() as i64,
                "w" => (date.weekday().num_days_from_sunday() + 1) as i64,
                "ww" => date.iso_week().week() as i64,
                _ => 0
            };
            
            Ok(Some(Value::Integer(part)))
        }

        // ============================================================
        // FORMATTING
        // ============================================================

        // FORMATDATETIME — FormatDateTime(date, [namedformat])
        "formatdatetime" => {
            if args.is_empty() {
                return Ok(Some(Value::String(String::new())));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let fmt = if args.len() > 1 {
                match evaluate_expression(&args[1], ctx)? {
                    Value::Integer(i) => i,
                    _ => 0
                }
            } else { 0 };
            
            if let Value::Date(d) = val {
                let result = match fmt {
                    0 => d.format("%m/%d/%Y").to_string(),   // vbGeneralDate
                    1 => d.format("%B %d, %Y").to_string(),  // vbLongDate
                    2 => d.format("%m/%d/%Y").to_string(),   // vbShortDate
                    3 => "00:00:00".to_string(),             // vbLongTime (no time in Date)
                    4 => "00:00".to_string(),                // vbShortTime
                    _ => d.format("%m/%d/%Y").to_string()
                };
                Ok(Some(Value::String(result)))
            } else {
                Ok(Some(Value::String(String::new())))
            }
        }

        // MONTHNAME — MonthName(month, [abbreviate])
        "monthname" => {
            if args.is_empty() {
                return Ok(Some(Value::String(String::new())));
            }
            let month_val = evaluate_expression(&args[0], ctx)?;
            let abbreviate = if args.len() > 1 {
                match evaluate_expression(&args[1], ctx)? {
                    Value::Boolean(b) => b,
                    _ => false
                }
            } else { false };
            
            let month = match month_val {
                Value::Integer(i) => i,
                _ => return Ok(Some(Value::String(String::new())))
            };
            
            let names_full = ["January", "February", "March", "April", "May", "June",
                             "July", "August", "September", "October", "November", "December"];
            let names_abbrev = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            
            if month >= 1 && month <= 12 {
                let name = if abbreviate {
                    names_abbrev[(month - 1) as usize]
                } else {
                    names_full[(month - 1) as usize]
                };
                Ok(Some(Value::String(name.to_string())))
            } else {
                Ok(Some(Value::String(String::new())))
            }
        }

        // ============================================================
        // TYPE CHECKING
        // ============================================================

        // ISDATE — Returns True if expression can be converted to a date
        "isdate" => {
            if args.is_empty() {
                return Ok(Some(Value::Boolean(false)));
            }
            let val = evaluate_expression(&args[0], ctx)?;
            let is_date = match val {
                Value::Date(_) => true,
                Value::DateTime(_) => true,
                Value::Time(_) => true,  // Time is also considered a date value in VBA
                Value::String(s) => {
                    let formats = ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y"];
                    formats.iter().any(|fmt| NaiveDate::parse_from_str(&s, fmt).is_ok())
                }
                _ => false
            };
            Ok(Some(Value::Boolean(is_date)))
        }

        _ => Ok(None)
    }
}
