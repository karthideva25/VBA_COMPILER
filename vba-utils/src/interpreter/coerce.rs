use anyhow::{anyhow, bail, Result};
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
        Value::Double(f)  => *f != 0.0,
        Value::Decimal(f) => *f != 0.0,
        Value::Single(f) => *f != 0.0,                 // NEW
        Value::Object(None) => false,                  // Nothing -> False
        Value::Object(Some(inner)) => to_bool(inner)?, // delegate
        Value::String(s)  => {
            let t = s.trim();
            if t.is_empty() { false }        // "" → False
            else { parse_bool(t)? }
        },
        Value::Array(_) => true, 
        Value::UserType { type_name, .. } => { // ADD THIS
            bail!("Cannot convert {} to Boolean", type_name)
        }
    })
}
pub(crate) fn to_i64(v: &Value) -> Result<i64> {
    Ok(match v {
        Value::Integer(i) => *i,
        Value::Long(l) => *l as i64,
        Value::LongLong(ll) => *ll,
        Value::Byte(b)    => *b as i64,
        Value::Currency(c) => *c as i64,
        Value::Date(d) => {
            // VBA stores dates as floating-point OLE Automation Dates
            // Integer part = days since 1899-12-30
            let base = chrono::NaiveDate::from_ymd_opt(1899, 12, 30)
                .ok_or_else(|| anyhow::anyhow!("invalid base date"))?;
            (d.signed_duration_since(base)).num_days()
        },
        Value::Boolean(b) => if *b { -1 } else { 0 },
        Value::Double(f)  => *f as i64,
        Value::Decimal(f) => *f as i64,
        Value::Single(f) => *f as i64,            // NEW
        Value::Object(Some(inner)) => to_i64(inner)?, // delegate if inner numeric
        Value::Object(None) => Err(anyhow!("Cannot convert Nothing to integer"))?,
        Value::String(s)  => {
            let t = s.trim();
            if t.is_empty() { 0 }            // "" → 0
            else { t.parse::<f64>().map_err(|e| anyhow!(e))? as i64 }
        },
        Value::UserType { type_name, .. } => { 
            bail!("Cannot convert {} to Integer", type_name)
        }
        &Value::Array(_) => bail!("Cannot convert Array to Integer")
    })
}
pub(crate) fn to_f64(v: &Value) -> Result<f64> {
    Ok(match v {
        Value::Integer(i) => *i as f64,
        Value::Long(l) => *l as f64,
        Value::LongLong(ll) => *ll as f64,
        Value::Byte(b)    => *b as f64,
        Value::Boolean(b) => if *b { -1.0 } else { 0.0 },
        Value::Currency(c) => *c,
        Value::Date(_) => 0.0, // or serialize to OLE Automation date if needed
        Value::Double(f)  => *f,
        Value::Decimal(f) => *f,
        Value::Single(f) => *f as f64,           
        Value::Object(Some(inner)) => to_f64(inner)?, // delegate
        Value::Object(None) => Err(anyhow!("Cannot convert Nothing to number"))?,
        Value::String(s)  => {
            let t = s.trim();
            if t.is_empty() { 0.0 }          // "" → 0.0
            else { t.parse::<f64>().map_err(|e| anyhow!(e))? }
        },
        Value::UserType { type_name, .. } => {
            bail!("Cannot convert {} to Double", type_name)
        }
        &Value::Array(_) => bail!("Cannot convert Array to Double")
    })
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
        Value::String(s)  => s.clone(),
        Value::Double(f)  => f.to_string(),
        Value::Decimal(f) => f.to_string(),
        Value::Single(f) => f.to_string(),                   
        Value::Object(None) => "Nothing".into(),             
        Value::Object(Some(inner)) => to_string(inner), 
        Value::Array(arr) => format!("Array({})", arr.len()),      
        Value::UserType { type_name, .. } => { 
            format!("<{} instance>", type_name)
        }
    }
}
pub(crate) fn add(l: Value, r: Value) -> Result<Value> {
    Ok(match (l, r) {
        (Value::Currency(a), Value::Currency(b)) => Value::Currency(a + b),
        (Value::Currency(a), b) => Value::Currency(a + to_f64(&b)?),
        (a, Value::Currency(b)) => Value::Currency(to_f64(&a)? + b),
        (Value::String(a), Value::String(b)) => Value::String(a + &b),
        (a @ Value::String(_), b) => Value::String(to_string(&a) + &to_string(&b)),
        (a, b @ Value::String(_)) => Value::String(to_string(&a) + &to_string(&b)),
        (a, b) => {
            // promote to floating point if any operand is Double/Decimal
            if matches!(a, Value::Double(_) | Value::Decimal(_) | Value::Single(_))
               || matches!(b, Value::Double(_) | Value::Decimal(_) | Value::Single(_)) {
            Value::Double(to_f64(&a)? + to_f64(&b)?)
        } else {
            Value::Integer(to_i64(&a)? + to_i64(&b)?)
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
            // Convert to integer first, then range-check 0..=255
            let n = to_i64(&val)?;
            if (0..=255).contains(&n) {
                Ok(Value::Byte(n as u8))
            } else {
                Err(anyhow!("overflow: {} does not fit in Byte (0..=255)", n))
            }
        }

        DT::Integer => {
            // VBA: True => -1, False => 0 is already handled by to_i64()
            let n = to_i64(&val)?;
            Ok(Value::Integer(n))
        }

        DT::Long => {
        // convert to i64 first then fit into i32 range
        let n = to_i64(&val)?;
        if (i32::MIN as i64 ..= i32::MAX as i64).contains(&n) {
            Ok(Value::Long(n as i32))
        } else {
            Err(anyhow!("overflow: {} does not fit in Long (i32)", n))
        }
    }

    DT::LongLong => {
        // LongLong is 64-bit signed
        let n = to_i64(&val)?;
        Ok(Value::LongLong(n))
    }

    DT::Object => {
        // If already an Object, keep it. Otherwise box the value as a reference.
        match val {
            Value::Object(_) => Ok(val),
            other => Ok(Value::Object(Some(Box::new(other)))),
        }
    }

    DT::Single => {
        let f = to_f64(&val)?;
        // store as 32-bit Single
        Ok(Value::Single(f as f32))
    }
    
    DT::String => Ok(Value::String(to_string(&val))),

        DT::Boolean => {
            let b = to_bool(&val)?;
            Ok(Value::Boolean(b))
        }

        DT::Currency => {
            // Keep as f64 for now; if you later switch to fixed-point/rounded, adjust here.
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

        DT::Date => {
            match val {
                Value::Date(d) => Ok(Value::Date(d)),
                Value::String(s) => {
                    let parsed = chrono::NaiveDate::parse_from_str(s.trim(), "%m/%d/%Y")
                        .map_err(|_| anyhow!("cannot parse '{}' as Date (mm/dd/yyyy)", s))?;
                    Ok(Value::Date(parsed))
                }
                // You could convert numerics via OLE Automation date here if you like.
                other => Err(anyhow!("cannot assign {:?} to Date", other)),
            }
        }

        DT::Variant => Ok(val),
    }
}

