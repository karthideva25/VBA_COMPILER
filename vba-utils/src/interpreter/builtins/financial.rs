// VBA Financial Functions
// This module implements VBA financial functions for financial calculations
// including depreciation, present/future value, payments, and rate of return.

use anyhow::Result;

use crate::context::{Context, Value};
use crate::ast::Expression;
use crate::interpreter::evaluate_expression;

/// Main handler for financial function calls
pub fn handle_financial_function(
    name: &str,
    args: &[Expression],
    ctx: &mut Context,
) -> Result<Option<Value>> {
    match name {
        // ============================================================
        // DEPRECIATION FUNCTIONS
        // ============================================================

        // SLN — Straight-line depreciation
        // SLN(Cost, Salvage, Life)
        "sln" => {
            if args.len() < 3 {
                return Ok(Some(Value::Double(0.0)));
            }
            let cost = get_numeric_arg(&args[0], ctx)?;
            let salvage = get_numeric_arg(&args[1], ctx)?;
            let life = get_numeric_arg(&args[2], ctx)?;
            
            if life == 0.0 {
                return Ok(Some(Value::Double(0.0)));
            }
            
            // Formula: (Cost - Salvage) / Life
            let result = (cost - salvage) / life;
            Ok(Some(Value::Double(result)))
        }

        // SYD — Sum-of-years-digits depreciation
        // SYD(Cost, Salvage, Life, Period)
        "syd" => {
            if args.len() < 4 {
                return Ok(Some(Value::Double(0.0)));
            }
            let cost = get_numeric_arg(&args[0], ctx)?;
            let salvage = get_numeric_arg(&args[1], ctx)?;
            let life = get_numeric_arg(&args[2], ctx)?;
            let period = get_numeric_arg(&args[3], ctx)?;
            
            if life <= 0.0 || period <= 0.0 || period > life {
                return Ok(Some(Value::Double(0.0)));
            }
            
            // Formula: (Cost - Salvage) * (Life - Period + 1) * 2 / (Life * (Life + 1))
            let sum_of_years = life * (life + 1.0) / 2.0;
            let remaining_life = life - period + 1.0;
            let result = (cost - salvage) * remaining_life / sum_of_years;
            Ok(Some(Value::Double(result)))
        }

        // DDB — Double-declining balance depreciation
        // DDB(Cost, Salvage, Life, Period, [Factor])
        "ddb" => {
            if args.len() < 4 {
                return Ok(Some(Value::Double(0.0)));
            }
            let cost = get_numeric_arg(&args[0], ctx)?;
            let salvage = get_numeric_arg(&args[1], ctx)?;
            let life = get_numeric_arg(&args[2], ctx)?;
            let period = get_numeric_arg(&args[3], ctx)?;
            let factor = if args.len() > 4 {
                get_numeric_arg(&args[4], ctx)?
            } else {
                2.0  // Default is double-declining (factor = 2)
            };
            
            if life <= 0.0 || period <= 0.0 || period > life {
                return Ok(Some(Value::Double(0.0)));
            }
            
            // Calculate accumulated depreciation up to previous period
            let rate = factor / life;
            let mut book_value = cost;
            
            for _ in 1..period as i64 {
                let depreciation = book_value * rate;
                book_value -= depreciation;
                if book_value < salvage {
                    book_value = salvage;
                    break;
                }
            }
            
            // Depreciation for the requested period
            let mut result = book_value * rate;
            if book_value - result < salvage {
                result = book_value - salvage;
            }
            if result < 0.0 {
                result = 0.0;
            }
            
            Ok(Some(Value::Double(result)))
        }

        // ============================================================
        // PRESENT/FUTURE VALUE FUNCTIONS
        // ============================================================

        // FV — Future value of an investment
        // FV(Rate, NPer, Pmt, [PV], [Due])
        "fv" => {
            if args.len() < 3 {
                return Ok(Some(Value::Double(0.0)));
            }
            let rate = get_numeric_arg(&args[0], ctx)?;
            let nper = get_numeric_arg(&args[1], ctx)?;
            let pmt = get_numeric_arg(&args[2], ctx)?;
            let pv = if args.len() > 3 { get_numeric_arg(&args[3], ctx)? } else { 0.0 };
            let due = if args.len() > 4 { get_numeric_arg(&args[4], ctx)? as i64 } else { 0 };
            
            let result = if rate == 0.0 {
                // Simple case: no interest
                -(pv + pmt * nper)
            } else {
                let factor = (1.0 + rate).powf(nper);
                let annuity_factor = if due != 0 {
                    // Payments at beginning of period
                    (factor - 1.0) / rate * (1.0 + rate)
                } else {
                    // Payments at end of period
                    (factor - 1.0) / rate
                };
                -(pv * factor + pmt * annuity_factor)
            };
            
            Ok(Some(Value::Double(result)))
        }

        // PV — Present value of an investment
        // PV(Rate, NPer, Pmt, [FV], [Due])
        "pv" => {
            if args.len() < 3 {
                return Ok(Some(Value::Double(0.0)));
            }
            let rate = get_numeric_arg(&args[0], ctx)?;
            let nper = get_numeric_arg(&args[1], ctx)?;
            let pmt = get_numeric_arg(&args[2], ctx)?;
            let fv = if args.len() > 3 { get_numeric_arg(&args[3], ctx)? } else { 0.0 };
            let due = if args.len() > 4 { get_numeric_arg(&args[4], ctx)? as i64 } else { 0 };
            
            let result = if rate == 0.0 {
                -(fv + pmt * nper)
            } else {
                let factor = (1.0 + rate).powf(nper);
                let annuity_factor = if due != 0 {
                    (1.0 - 1.0 / factor) / rate * (1.0 + rate)
                } else {
                    (1.0 - 1.0 / factor) / rate
                };
                -(fv / factor + pmt * annuity_factor)
            };
            
            Ok(Some(Value::Double(result)))
        }

        // NPV — Net present value of cash flows
        // NPV(Rate, Value1, Value2, ...)
        "npv" => {
            if args.len() < 2 {
                return Ok(Some(Value::Double(0.0)));
            }
            let rate = get_numeric_arg(&args[0], ctx)?;
            
            // Collect all cash flows from remaining arguments
            let mut npv = 0.0;
            let mut period = 1;
            
            for i in 1..args.len() {
                let val = evaluate_expression(&args[i], ctx)?;
                let cf = value_to_f64(&val);
                npv += cf / (1.0 + rate).powi(period);
                period += 1;
            }
            
            Ok(Some(Value::Double(npv)))
        }

        // ============================================================
        // PAYMENT FUNCTIONS
        // ============================================================

        // PMT — Payment for a loan
        // Pmt(Rate, NPer, PV, [FV], [Due])
        "pmt" => {
            if args.len() < 3 {
                return Ok(Some(Value::Double(0.0)));
            }
            let rate = get_numeric_arg(&args[0], ctx)?;
            let nper = get_numeric_arg(&args[1], ctx)?;
            let pv = get_numeric_arg(&args[2], ctx)?;
            let fv = if args.len() > 3 { get_numeric_arg(&args[3], ctx)? } else { 0.0 };
            let due = if args.len() > 4 { get_numeric_arg(&args[4], ctx)? as i64 } else { 0 };
            
            let result = if rate == 0.0 {
                -(pv + fv) / nper
            } else {
                let factor = (1.0 + rate).powf(nper);
                let pmt = if due != 0 {
                    // Payments at beginning of period
                    (rate * (pv * factor + fv)) / ((factor - 1.0) * (1.0 + rate))
                } else {
                    // Payments at end of period
                    (rate * (pv * factor + fv)) / (factor - 1.0)
                };
                -pmt
            };
            
            Ok(Some(Value::Double(result)))
        }

        // IPMT — Interest payment for a specific period
        // IPmt(Rate, Per, NPer, PV, [FV], [Due])
        "ipmt" => {
            if args.len() < 4 {
                return Ok(Some(Value::Double(0.0)));
            }
            let rate = get_numeric_arg(&args[0], ctx)?;
            let per = get_numeric_arg(&args[1], ctx)?;
            let nper = get_numeric_arg(&args[2], ctx)?;
            let pv = get_numeric_arg(&args[3], ctx)?;
            let fv = if args.len() > 4 { get_numeric_arg(&args[4], ctx)? } else { 0.0 };
            let due = if args.len() > 5 { get_numeric_arg(&args[5], ctx)? as i64 } else { 0 };
            
            if per < 1.0 || per > nper {
                return Ok(Some(Value::Double(0.0)));
            }
            
            // Calculate PMT first
            let pmt = if rate == 0.0 {
                -(pv + fv) / nper
            } else {
                let factor = (1.0 + rate).powf(nper);
                if due != 0 {
                    -(rate * (pv * factor + fv)) / ((factor - 1.0) * (1.0 + rate))
                } else {
                    -(rate * (pv * factor + fv)) / (factor - 1.0)
                }
            };
            
            // Calculate balance at start of period
            let balance = if rate == 0.0 {
                pv + pmt * (per - 1.0)
            } else {
                let factor = (1.0 + rate).powf(per - 1.0);
                pv * factor + pmt * (factor - 1.0) / rate
            };
            
            // Interest for this period
            let result = if due != 0 && per == 1.0 {
                0.0  // No interest on first payment if due at beginning
            } else if due != 0 {
                (balance - pmt) * rate
            } else {
                balance * rate
            };
            
            Ok(Some(Value::Double(result)))
        }

        // PPMT — Principal payment for a specific period
        // PPmt(Rate, Per, NPer, PV, [FV], [Due])
        "ppmt" => {
            if args.len() < 4 {
                return Ok(Some(Value::Double(0.0)));
            }
            let rate = get_numeric_arg(&args[0], ctx)?;
            let per = get_numeric_arg(&args[1], ctx)?;
            let nper = get_numeric_arg(&args[2], ctx)?;
            let pv = get_numeric_arg(&args[3], ctx)?;
            let fv = if args.len() > 4 { get_numeric_arg(&args[4], ctx)? } else { 0.0 };
            let due = if args.len() > 5 { get_numeric_arg(&args[5], ctx)? as i64 } else { 0 };
            
            if per < 1.0 || per > nper {
                return Ok(Some(Value::Double(0.0)));
            }
            
            // PMT = PPMT + IPMT, so PPMT = PMT - IPMT
            // Calculate PMT
            let pmt = if rate == 0.0 {
                -(pv + fv) / nper
            } else {
                let factor = (1.0 + rate).powf(nper);
                if due != 0 {
                    -(rate * (pv * factor + fv)) / ((factor - 1.0) * (1.0 + rate))
                } else {
                    -(rate * (pv * factor + fv)) / (factor - 1.0)
                }
            };
            
            // Calculate IPMT
            let balance = if rate == 0.0 {
                pv + pmt * (per - 1.0)
            } else {
                let factor = (1.0 + rate).powf(per - 1.0);
                pv * factor + pmt * (factor - 1.0) / rate
            };
            
            let ipmt = if due != 0 && per == 1.0 {
                0.0
            } else if due != 0 {
                (balance - pmt) * rate
            } else {
                balance * rate
            };
            
            // PPMT = PMT - IPMT
            Ok(Some(Value::Double(pmt - ipmt)))
        }

        // ============================================================
        // LOAN/INVESTMENT FUNCTIONS
        // ============================================================

        // NPER — Number of periods for an investment
        // NPer(Rate, Pmt, PV, [FV], [Due])
        "nper" => {
            if args.len() < 3 {
                return Ok(Some(Value::Double(0.0)));
            }
            let rate = get_numeric_arg(&args[0], ctx)?;
            let pmt = get_numeric_arg(&args[1], ctx)?;
            let pv = get_numeric_arg(&args[2], ctx)?;
            let fv = if args.len() > 3 { get_numeric_arg(&args[3], ctx)? } else { 0.0 };
            let due = if args.len() > 4 { get_numeric_arg(&args[4], ctx)? as i64 } else { 0 };
            
            let result = if rate.abs() < 1e-10 {
                // Zero interest rate case
                if pmt.abs() < 1e-10 {
                    f64::NAN  // Can't solve with zero payment and zero rate
                } else {
                    -(pv + fv) / pmt
                }
            } else {
                // Standard NPER formula
                let type_factor = if due != 0 { 1.0 + rate } else { 1.0 };
                let pmt_adj = pmt * type_factor;
                
                let a = pmt_adj / rate;
                let numerator = a - fv;
                let denominator = a + pv;
                
                if denominator.abs() < 1e-10 || numerator / denominator <= 0.0 {
                    f64::NAN
                } else {
                    (numerator / denominator).ln() / (1.0 + rate).ln()
                }
            };
            
            Ok(Some(Value::Double(result)))
        }

        // RATE — Interest rate per period
        // Rate(NPer, Pmt, PV, [FV], [Due], [Guess])
        "rate" => {
            if args.len() < 3 {
                return Ok(Some(Value::Double(0.0)));
            }
            let nper = get_numeric_arg(&args[0], ctx)?;
            let pmt = get_numeric_arg(&args[1], ctx)?;
            let pv = get_numeric_arg(&args[2], ctx)?;
            let fv = if args.len() > 3 { get_numeric_arg(&args[3], ctx)? } else { 0.0 };
            let due = if args.len() > 4 { get_numeric_arg(&args[4], ctx)? as i64 } else { 0 };
            let guess = if args.len() > 5 { get_numeric_arg(&args[5], ctx)? } else { 0.1 };
            
            // Use Newton-Raphson iteration to find rate
            let mut rate = guess;
            let max_iterations = 100;
            let tolerance = 1e-10;
            
            for _ in 0..max_iterations {
                let factor = (1.0 + rate).powf(nper);
                let pmt_factor = if due != 0 { 1.0 + rate } else { 1.0 };
                
                // f(rate) = pv * factor + pmt * pmt_factor * (factor - 1) / rate + fv
                let f = if rate == 0.0 {
                    pv + pmt * nper + fv
                } else {
                    pv * factor + pmt * pmt_factor * (factor - 1.0) / rate + fv
                };
                
                // f'(rate) - derivative
                let df = if rate.abs() < 1e-10 {
                    pv * nper + pmt * nper * (nper - 1.0) / 2.0
                } else {
                    let factor_deriv = nper * (1.0 + rate).powf(nper - 1.0);
                    pv * factor_deriv + 
                    pmt * pmt_factor * (rate * factor_deriv - (factor - 1.0)) / (rate * rate) +
                    if due != 0 { pmt * (factor - 1.0) / rate } else { 0.0 }
                };
                
                if df.abs() < 1e-20 {
                    break;
                }
                
                let new_rate = rate - f / df;
                
                if (new_rate - rate).abs() < tolerance {
                    rate = new_rate;
                    break;
                }
                rate = new_rate;
            }
            
            Ok(Some(Value::Double(rate)))
        }

        // ============================================================
        // INTERNAL RATE OF RETURN FUNCTIONS
        // ============================================================

        // IRR — Internal rate of return for cash flows
        // IRR(CashFlow1, CashFlow2, ..., [Guess])
        "irr" => {
            if args.len() < 2 {
                return Ok(Some(Value::Double(0.0)));
            }
            
            // Collect cash flows from arguments
            // Last argument might be a guess if it's very small (0.0-1.0 range)
            let mut cash_flows: Vec<f64> = Vec::new();
            let mut guess = 0.1;
            
            for (i, arg) in args.iter().enumerate() {
                let v = evaluate_expression(arg, ctx)?;
                let val = value_to_f64(&v);
                
                // Check if last arg might be a guess (typically 0.0 to 1.0)
                if i == args.len() - 1 && args.len() > 2 && val.abs() <= 1.0 && val > -1.0 {
                    guess = val;
                } else {
                    cash_flows.push(val);
                }
            }
            
            if cash_flows.len() < 2 {
                return Ok(Some(Value::Double(0.0)));
            }
            
            // Newton-Raphson iteration for IRR
            let mut rate = guess;
            let max_iterations = 100;
            let tolerance = 1e-10;
            
            for _ in 0..max_iterations {
                let mut npv = 0.0;
                let mut npv_deriv = 0.0;
                
                for (t, cf) in cash_flows.iter().enumerate() {
                    let factor = (1.0 + rate).powi(t as i32);
                    npv += cf / factor;
                    if t > 0 {
                        npv_deriv -= (t as f64) * cf / ((1.0 + rate).powi(t as i32 + 1));
                    }
                }
                
                if npv_deriv.abs() < 1e-20 {
                    break;
                }
                
                let new_rate = rate - npv / npv_deriv;
                
                if (new_rate - rate).abs() < tolerance {
                    rate = new_rate;
                    break;
                }
                rate = new_rate;
            }
            
            Ok(Some(Value::Double(rate)))
        }

        // MIRR — Modified internal rate of return
        // MIRR(CashFlow1, CashFlow2, ..., FinanceRate, ReinvestRate)
        "mirr" => {
            if args.len() < 4 {
                // Need at least 2 cash flows + 2 rates
                return Ok(Some(Value::Double(0.0)));
            }
            
            // Last two args are rates
            let reinvest_rate = get_numeric_arg(&args[args.len() - 1], ctx)?;
            let finance_rate = get_numeric_arg(&args[args.len() - 2], ctx)?;
            
            // Collect cash flows (all args except last two)
            let mut cash_flows: Vec<f64> = Vec::new();
            for i in 0..(args.len() - 2) {
                let val = get_numeric_arg(&args[i], ctx)?;
                cash_flows.push(val);
            }
            
            if cash_flows.is_empty() {
                return Ok(Some(Value::Double(0.0)));
            }
            
            let n = cash_flows.len() as i32;
            
            // Calculate NPV of negative cash flows (costs) at finance rate
            let mut npv_neg = 0.0;
            for (t, cf) in cash_flows.iter().enumerate() {
                if *cf < 0.0 {
                    npv_neg += cf / (1.0 + finance_rate).powi(t as i32);
                }
            }
            
            // Calculate FV of positive cash flows at reinvest rate
            let mut fv_pos = 0.0;
            for (t, cf) in cash_flows.iter().enumerate() {
                if *cf > 0.0 {
                    fv_pos += cf * (1.0 + reinvest_rate).powi(n - 1 - t as i32);
                }
            }
            
            // MIRR formula
            let result = if npv_neg == 0.0 || fv_pos == 0.0 {
                0.0
            } else {
                (fv_pos / (-npv_neg)).powf(1.0 / (n as f64 - 1.0)) - 1.0
            };
            
            Ok(Some(Value::Double(result)))
        }

        _ => Ok(None)
    }
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================

/// Extract numeric value from an expression
fn get_numeric_arg(expr: &Expression, ctx: &mut Context) -> Result<f64> {
    let val = evaluate_expression(expr, ctx)?;
    Ok(value_to_f64(&val))
}

/// Convert a Value to f64
pub fn value_to_f64(val: &Value) -> f64 {
    match val {
        Value::Integer(i) => *i as f64,
        Value::Long(l) => *l as f64,
        Value::LongLong(ll) => *ll as f64,
        Value::Double(d) => *d,
        Value::Single(s) => *s as f64,
        Value::Currency(c) => *c,
        Value::Byte(b) => *b as f64,
        _ => 0.0,
    }
}
