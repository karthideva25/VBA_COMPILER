use anyhow::{anyhow, Result};
use crate::context::{Context, ErrObject, Value};
use super::coerce;

// Small helper
fn set_err(ctx: &mut Context, number: i32, description: &str) {
    ctx.err = Some(ErrObject {
        number,
        description: description.to_string(),
        source: "Interpreter".into(),
    });
}

// Unary stays pure (no ctx needed)
pub(crate) fn eval_unary(op: &str, v: Value) -> Result<Value> {
    match op {
        "+" => Ok(Value::Double(super::coerce::to_f64(&v)?)),
        "-" => Ok(Value::Double(-super::coerce::to_f64(&v)?)),
        "Not" | "not" | "NOT" => Ok(Value::Boolean(!super::coerce::to_bool(&v)?)),
        other => Err(anyhow!("unary op not implemented: {}", other)),
    }
}

/// IMPORTANT: now takes &mut Context so we can set Err on runtime faults
pub(crate) fn eval_binary(ctx: &mut Context, op: &str, l: Value, r: Value) -> Result<Value> {
    match op {
        // VBA `+` is numeric add unless either side is a string (then concat via + rules).
        "+" => super::coerce::add(l, r),

        // `&` is *always* string concatenation in VBA
        "&" => {
            let ls = coerce::to_string(&l);
            let rs = coerce::to_string(&r);
            Ok(Value::String(ls + &rs))
        }

        "-" => {
            // Use checked arithmetic on integer-compatible values; fall back to Double as needed
            // VBA promotes as needed; a practical compromise: try i64 first, if conversion fails, do Double
            if let (Ok(li), Ok(ri)) = (coerce::to_i64(&l), coerce::to_i64(&r)) {
                match li.checked_sub(ri) {
                    Some(v) => {
                        //println!("✅ Subtraction successful: {} - {} = {}", li, ri, v);
                        Ok(Value::Integer(v))}
                    None => {
                        // Overflow → Err 6
                        //println!("🔴 OVERFLOW DETECTED in subtraction: {} - {}", li, ri);
                        set_err(ctx, 6, "Overflow");
                        Ok(Value::Integer(0)) // placeholder; Assignment guard will skip the write
                    }
                }
            } else {
                let result = coerce::to_f64(&l)? - coerce::to_f64(&r)?;
                //println!("✅ Subtraction (as double): result = {}", result);
                Ok(Value::Double(coerce::to_f64(&l)? - coerce::to_f64(&r)?))
            }
        }

        "*" => {
            if let (Ok(li), Ok(ri)) = (coerce::to_i64(&l), coerce::to_i64(&r)) {
                match li.checked_mul(ri) {
                    Some(v) => Ok(Value::Integer(v)),
                    None => {
                        set_err(ctx, 6, "Overflow");
                        Ok(Value::Integer(0))
                    }
                }
            } else {
                Ok(Value::Double(coerce::to_f64(&l)? * coerce::to_f64(&r)?))
            }
        }

        "/" => {
            let denom = super::coerce::to_f64(&r)?;
            if denom == 0.0 {
                // This bubbles as Err up to Assignment where we set Err.Number = 11
                //println!("🔴 DIVISION BY ZERO DETECTED! Setting error...");
                set_err(ctx, 11, "Division by zero");
                return Ok(Value::Double(f64::NAN));
            }
            let num = super::coerce::to_f64(&l)?;
            let result = num / denom;
            //println!("✅ Division successful: {} / {} = {}", num, denom, result);
            Ok(Value::Double(num / denom))
        }

        // Logical
        "And" | "and" => Ok(Value::Boolean(coerce::to_bool(&l)? & coerce::to_bool(&r)?)),
        "Or"  | "or"  => Ok(Value::Boolean(coerce::to_bool(&l)? | coerce::to_bool(&r)?)),

        // Comparisons: coerce to VBA-like numeric comparison for non-strings
        "="  => Ok(Value::Boolean(coerce::cmp_eq(&l, &r)?)),
        "<>" => Ok(Value::Boolean(!coerce::cmp_eq(&l, &r)?)),
        "<"  => Ok(Value::Boolean(coerce::to_f64(&l)? <  coerce::to_f64(&r)?)),
        "<=" => Ok(Value::Boolean(coerce::to_f64(&l)? <= coerce::to_f64(&r)?)),
        ">"  => Ok(Value::Boolean(coerce::to_f64(&l)? >  coerce::to_f64(&r)?)),
        ">=" => Ok(Value::Boolean(coerce::to_f64(&l)? >= coerce::to_f64(&r)?)),

        other => Err(anyhow!("binary op not implemented: {}", other)),
    }
}
