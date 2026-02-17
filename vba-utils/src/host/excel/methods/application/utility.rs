// Utility and conversion methods

use anyhow::Result;
use crate::context::Value;

pub fn call_method(method: &str, args: &[Value]) -> Result<Value> {
    match method.to_lowercase().as_str() {
        "centimeterstopo" => {
            // Convert centimeters to points
            if let Some(Value::Double(cm)) = args.first() {
                let points = cm * 28.346456693; // 1 cm = 28.346... points
                eprintln!("📐 CentimetersToPoints({}) = {}", cm, points);
                Ok(Value::Double(points))
            } else {
                Err(anyhow::anyhow!("CentimetersToPoints requires a numeric argument"))
            }
        }
        "inchestopoint" => {
            // Convert inches to points
            if let Some(Value::Double(inches)) = args.first() {
                let points = inches * 72.0; // 1 inch = 72 points
                eprintln!("📏 InchesToPoints({}) = {}", inches, points);
                Ok(Value::Double(points))
            } else {
                Err(anyhow::anyhow!("InchesToPoints requires a numeric argument"))
            }
        }
        "convertformula" => {
            eprintln!("🔄 Application.ConvertFormula() - converting formula between styles");
            Ok(Value::String(String::new()))
        }
        "evaluate" => {
            eprintln!("🔢 Application.Evaluate() - evaluating expression");
            Ok(Value::Empty)
        }
        _ => Err(anyhow::anyhow!("Unknown utility method: {}", method)),
    }
}
