use anyhow::{anyhow, bail, Result};
use chrono::Timelike;  // For hour(), minute(), second() on NaiveTime
use crate::context::{Value, DeclaredType};

pub(crate) fn to_bool(v: &Value) -> Result<bool> {
    Ok(match v {
        Value::Boolean(b) => *b,
        Value::Integer(i) => *i != 0,           // VBA True = -1 (non-zero)
        Value::Long(l) => *l != 0,
        Value::LongLong(ll) => *ll != 0,
        Value::Byte(b)    => *b != 0,
        Value::Currency(c) => *c != 0.0,        // Currency behaves like numeric
        Value::Date(_) => true,                 // Any valid date is True
        Value::DateTime(_) => true,             // Any valid datetime is True
        Value::Time(_) => true,                 // Any valid time is True
        Value::Double(f)  => *f != 0.0,
        Value::Decimal(f) => *f != 0.0,
        Value::Single(f) => *f != 0.0,
        Value::Object(None) => false,                  // Nothing -> False
        Value::Object(Some(inner)) => to_bool(inner)?, // delegate
        Value::String(s)  => {
            let t = s.trim();
            if t.is_empty() { false }        // "" → False
            else { parse_bool(t)? }
        },
        Value::UserType { type_name, .. } => {
            bail!("Cannot convert {} to Boolean", type_name)
        }
        Value::Empty => false,
        Value::Null => false,
        Value::Error(_) => bail!("Cannot convert Error to Boolean"),
    })
}

pub(crate) fn to_i64(v: &Value) -> Result<i64> {
    use Value::*;

    match v {
        Integer(i) => Ok(*i),
        Long(l) => Ok(*l as i64),
        LongLong(ll) => Ok(*ll),
        Byte(b) => Ok(*b as i64),
        Currency(c) => Ok(*c as i64),

        Date(d) => {
            // VBA stores dates as floating-point OLE Automation Dates
            // Integer part = days since 1899-12-30
            let base = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
                .ok_or_else(|| anyhow!("invalid base date"))?;
            Ok(d.signed_duration_since(base).num_days())
        }
        
        DateTime(dt) => {
            let base = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
                .ok_or_else(|| anyhow!("invalid base date"))?;
            Ok(dt.date().signed_duration_since(base).num_days())
        }
        
        Time(_) => Ok(0), // Time alone has no date component

        Boolean(b) => Ok(if *b { -1 } else { 0 }),
        Double(f) => Ok(*f as i64),
        Decimal(f) => Ok(*f as i64),
        Single(f) => Ok(*f as i64),

        Object(Some(inner)) => to_i64(inner),
        Object(None) => Err(anyhow!("Cannot convert Nothing to integer")),

        String(s) => {
            let t = s.trim();
            if t.is_empty() {
                Ok(0) // "" → 0
            } else {
                let f: f64 = t
                    .parse()
                    .map_err(|e| anyhow!("cannot parse '{t}' as number: {e}"))?;
                Ok(f as i64)
            }
        }

        UserType { type_name, .. } => {
            bail!("Cannot convert {} to Integer", type_name)
        }

        Empty => Ok(0),
        Null => Err(anyhow!("Cannot convert Null to integer")),
        Error(e) => Ok(*e as i64),  // Error values convert to their error number
    }
}

pub(crate) fn to_f64(v: &Value) -> Result<f64> {
    use Value::*;

    match v {
        Integer(i) => Ok(*i as f64),
        Long(l) => Ok(*l as f64),
        LongLong(ll) => Ok(*ll as f64),
        Byte(b) => Ok(*b as f64),

        Boolean(b) => Ok(if *b { -1.0 } else { 0.0 }),
        Currency(c) => Ok(*c),
        Date(_) => Ok(0.0), // or serialize to OLE Automation date if needed
        DateTime(_) => Ok(0.0),
        Time(_) => Ok(0.0),
        Double(f) => Ok(*f),
        Decimal(f) => Ok(*f),
        Single(f) => Ok(*f as f64),

        Object(Some(inner)) => to_f64(inner),
        Object(None) => Err(anyhow!("Cannot convert Nothing to number")),

        String(s) => {
            let t = s.trim();
            if t.is_empty() {
                Ok(0.0) // "" → 0.0
            } else {
                let f: f64 = t
                    .parse()
                    .map_err(|e| anyhow!("cannot parse '{t}' as number: {e}"))?;
                Ok(f)
            }
        }

        UserType { type_name, .. } => {
            bail!("Cannot convert {} to Double", type_name)
        }

        Empty => Ok(0.0),
        Null => Err(anyhow!("Cannot convert Null to Double")),
        Error(e) => Ok(*e as f64),  // Error values convert to their error number
    }
}

pub(crate) fn to_string(v: &Value) -> String {
    match v {
        Value::Integer(i) => i.to_string(),
        Value::Long(l) => l.to_string(),
        Value::LongLong(ll) => ll.to_string(),
        Value::Byte(b)    => b.to_string(),
        Value::Boolean(b) => if *b { "True" } else { "False" }.into(),
        Value::Currency(c) => format!("{:.4}", c),
        Value::Date(d) => d.format("%m/%d/%Y").to_string(),
        Value::DateTime(dt) => dt.format("%m/%d/%Y %H:%M:%S").to_string(),
        Value::Time(t) => t.format("%H:%M:%S").to_string(),
        Value::String(s)  => s.clone(),
        Value::Double(f)  => f.to_string(),
        Value::Decimal(f) => f.to_string(),
        Value::Single(f) => f.to_string(),
        Value::Object(None) => "Nothing".into(),
        Value::Object(Some(inner)) => to_string(inner),
        Value::UserType { type_name, .. } => {
            format!("<{} instance>", type_name)
        }
        Value::Empty => String::new(),
        Value::Null => String::new(), // Null becomes "" in string context
        Value::Error(e) => format!("Error {}", e),
    }
}

/// Helper to check if a string can be parsed as a number
fn is_numeric_string(s: &str) -> bool {
    s.trim().parse::<f64>().is_ok()
}

/// VBA + operator with Variant semantics:
/// - If both are strings that look like numbers, do numeric addition
/// - If both are strings that don't look like numbers, raise Type Mismatch (we concatenate as fallback)
/// - If one is numeric, try to coerce the other to numeric
/// - DateTime + Time = DateTime with time added
pub(crate) fn add(l: Value, r: Value) -> Result<Value> {
    // Handle Null propagation
    if matches!((&l, &r), (Value::Null, _) | (_, Value::Null)) {
        return Ok(Value::Null);
    }
    
    Ok(match (&l, &r) {
        // DateTime + Time arithmetic (e.g., Now() + TimeValue("00:00:10"))
        (Value::DateTime(dt), Value::Time(t)) => {
            // Add the time duration to the datetime
            let duration = chrono::Duration::hours(t.hour() as i64)
                + chrono::Duration::minutes(t.minute() as i64)
                + chrono::Duration::seconds(t.second() as i64);
            if let Some(new_dt) = dt.checked_add_signed(duration) {
                Value::DateTime(new_dt)
            } else {
                Value::DateTime(*dt)
            }
        }
        (Value::Time(t), Value::DateTime(dt)) => {
            // Time + DateTime = DateTime with time added
            let duration = chrono::Duration::hours(t.hour() as i64)
                + chrono::Duration::minutes(t.minute() as i64)
                + chrono::Duration::seconds(t.second() as i64);
            if let Some(new_dt) = dt.checked_add_signed(duration) {
                Value::DateTime(new_dt)
            } else {
                Value::DateTime(*dt)
            }
        }
        
        // Date + Time = DateTime
        (Value::Date(d), Value::Time(t)) => {
            let dt = d.and_time(*t);
            Value::DateTime(dt)
        }
        (Value::Time(t), Value::Date(d)) => {
            let dt = d.and_time(*t);
            Value::DateTime(dt)
        }
        
        // Time + Time = Time (add durations, wrap at 24h)
        (Value::Time(t1), Value::Time(t2)) => {
            // Calculate total seconds manually using Timelike trait
            let secs1 = t1.hour() * 3600 + t1.minute() * 60 + t1.second();
            let secs2 = t2.hour() * 3600 + t2.minute() * 60 + t2.second();
            let total_secs = secs1 + secs2;
            let wrapped_secs = total_secs % 86400; // 24 * 60 * 60
            if let Some(new_time) = chrono::NaiveTime::from_num_seconds_from_midnight_opt(wrapped_secs, 0) {
                Value::Time(new_time)
            } else {
                Value::Time(*t1)
            }
        }
        
        // Currency operations
        (Value::Currency(a), Value::Currency(b)) => Value::Currency(a + b),
        (Value::Currency(a), _) => Value::Currency(a + to_f64(&r)?),
        (_, Value::Currency(b)) => Value::Currency(to_f64(&l)? + b),
        
        // String + String: VBA tries numeric coercion with +
        (Value::String(a), Value::String(b)) => {
            // Both are strings - try numeric conversion
            if is_numeric_string(a) && is_numeric_string(b) {
                Value::Double(to_f64(&l)? + to_f64(&r)?)
            } else {
                // Type mismatch - in VBA this would be Error 13
                // For now, we'll concatenate as a fallback
                bail!("Type mismatch: cannot add non-numeric strings '{}' + '{}'", a, b)
            }
        }
        
        // String + Number or Number + String: try numeric coercion
        (Value::String(s), _) => {
            if is_numeric_string(s) {
                // promote to floating point based on other operand
                if matches!(r, Value::Double(_) | Value::Decimal(_) | Value::Single(_)) {
                    Value::Double(to_f64(&l)? + to_f64(&r)?)
                } else {
                    Value::Integer(to_i64(&l)? + to_i64(&r)?)
                }
            } else {
                bail!("Type mismatch: cannot convert '{}' to a number", s)
            }
        }
        (_, Value::String(s)) => {
            if is_numeric_string(s) {
                // promote to floating point based on other operand
                if matches!(l, Value::Double(_) | Value::Decimal(_) | Value::Single(_)) {
                    Value::Double(to_f64(&l)? + to_f64(&r)?)
                } else {
                    Value::Integer(to_i64(&l)? + to_i64(&r)?)
                }
            } else {
                bail!("Type mismatch: cannot convert '{}' to a number", s)
            }
        }
        
        // Numeric + Numeric
        _ => {
            // promote to floating point if any operand is Double/Decimal/Single
            if matches!(l, Value::Double(_) | Value::Decimal(_) | Value::Single(_))
                || matches!(r, Value::Double(_) | Value::Decimal(_) | Value::Single(_))
            {
                Value::Double(to_f64(&l)? + to_f64(&r)?)
            } else {
                Value::Integer(to_i64(&l)? + to_i64(&r)?)
            }
        }
    })
}

pub(crate) fn cmp_eq(l: &Value, r: &Value) -> Result<bool> {
    Ok(match (l, r) {
        (Value::String(a), Value::String(b)) => a == b,
        _ => (to_f64(l)? - to_f64(r)?).abs() < f64::EPSILON,
    })
}

/// Numeric comparison operators (<, <=, >, >=) can also reuse `to_f64`
pub(crate) fn cmp_lt(l: &Value, r: &Value) -> Result<bool> {
    Ok(to_f64(l)? < to_f64(r)?)
}

pub(crate) fn cmp_gt(l: &Value, r: &Value) -> Result<bool> {
    Ok(to_f64(l)? > to_f64(r)?)
}

fn parse_bool(s: &str) -> Result<bool> {
    match s.trim().to_ascii_lowercase().as_str() {
        "true" | "-1" => Ok(true),
        "false" | "0" => Ok(false),
        other => Err(anyhow!("cannot coerce string to boolean: {other}")),
    }
}

/// Coerce a runtime Value into the declared VBA type (used at assignment time).
/// This enforces Byte range, Boolean semantics (True = -1), basic Date parsing, etc.
pub(crate) fn coerce_to_declared(val: Value, ty: DeclaredType) -> Result<Value> {
    use DeclaredType as DT;

    match ty {
        DT::Byte => {
            let n = to_i64(&val)?;
            if (0..=255).contains(&n) {
                Ok(Value::Byte(n as u8))
            } else {
                Err(anyhow!("overflow: {} does not fit in Byte (0..=255)", n))
            }
        }

        DT::Integer => {
            let n = to_i64(&val)?;
            Ok(Value::Integer(n))
        }

        DT::Long => {
            let n = to_i64(&val)?;
            if (i32::MIN as i64..=i32::MAX as i64).contains(&n) {
                Ok(Value::Long(n as i32))
            } else {
                Err(anyhow!("overflow: {} does not fit in Long (i32)", n))
            }
        }

        DT::LongLong => {
            let n = to_i64(&val)?;
            Ok(Value::LongLong(n))
        }

        DT::Object => {
            match val {
                Value::Object(_) => Ok(val),
                other => Ok(Value::Object(Some(Box::new(other)))),
            }
        }

        DT::Single => {
            let f = to_f64(&val)?;
            Ok(Value::Single(f as f32))
        }

        DT::String => Ok(Value::String(to_string(&val))),

        DT::Boolean => {
            let b = to_bool(&val)?;
            Ok(Value::Boolean(b))
        }

        DT::Currency => {
            let f = to_f64(&val)?;
            Ok(Value::Currency(f))
        }

        DT::Double => {
            let f = to_f64(&val)?;
            Ok(Value::Double(f))
        }

        DT::Decimal => {
            let f = to_f64(&val)?;
            Ok(Value::Decimal(f))
        }

        DT::Date => match val {
            Value::Date(d) => Ok(Value::Date(d)),
            Value::String(s) => {
                let parsed = chrono::NaiveDate::parse_from_str(s.trim(), "%m/%d/%Y")
                    .map_err(|_| anyhow!("cannot parse '{}' as Date (mm/dd/yyyy)", s))?;
                Ok(Value::Date(parsed))
            }
            other => Err(anyhow!("cannot assign {:?} to Date", other)),
        },

        DT::Variant => Ok(val),
    }
}
