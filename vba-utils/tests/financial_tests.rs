// Tests for VBA Financial Functions
//
// This test file covers all 13 VBA financial functions:
// - Depreciation: SLN, SYD, DDB
// - Present/Future Value: FV, PV, NPV
// - Payment: Pmt, IPmt, PPmt
// - Loan/Investment: NPer, Rate
// - Internal Rate of Return: IRR, MIRR

use tree_sitter::Parser;
use vba_parser::language as tree_sitter_vba;
use vba_utils::Context;
use vba_utils::vm::ProgramExecutor;
use vba_utils::ast::build_ast;

/// Helper to run VBA code and capture output
/// Uses AutoOpen as the entry point
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
// DEPRECIATION FUNCTION TESTS
// ============================================================

#[test]
fn test_sln_basic() {
    // SLN(Cost, Salvage, Life) = (Cost - Salvage) / Life
    // (10000 - 1000) / 5 = 1800
    let code = r#"
        Sub AutoOpen()
            MsgBox SLN(10000, 1000, 5)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "1800");
}

#[test]
fn test_sln_zero_life() {
    // Life = 0 should return 0 (avoid division by zero)
    let code = r#"
        Sub AutoOpen()
            MsgBox SLN(10000, 1000, 0)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "0");
}

#[test]
fn test_syd_basic() {
    // SYD(Cost, Salvage, Life, Period)
    // For period 1 of 5 years: (10000-1000) * 5/15 = 3000
    let code = r#"
        Sub AutoOpen()
            MsgBox SYD(10000, 1000, 5, 1)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "3000");
}

#[test]
fn test_syd_period_2() {
    // Period 2: (10000-1000) * 4/15 = 2400
    let code = r#"
        Sub AutoOpen()
            MsgBox SYD(10000, 1000, 5, 2)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "2400");
}

#[test]
fn test_ddb_basic() {
    // DDB(Cost, Salvage, Life, Period, [Factor])
    // Default factor = 2 (double declining)
    // Period 1: 10000 * (2/5) = 4000
    let code = r#"
        Sub AutoOpen()
            MsgBox DDB(10000, 1000, 5, 1)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "4000");
}

#[test]
fn test_ddb_with_factor() {
    // DDB with factor = 1.5
    // Period 1: 10000 * (1.5/5) = 3000
    let code = r#"
        Sub AutoOpen()
            MsgBox DDB(10000, 1000, 5, 1, 1.5)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "3000");
}

// ============================================================
// PRESENT/FUTURE VALUE FUNCTION TESTS
// ============================================================

#[test]
fn test_fv_monthly_investment() {
    // FV(Rate, NPer, Pmt, [PV], [Due])
    // $100/month at 5% annual for 10 years
    // FV(0.05/12, 120, -100, 0, 0) ≈ 15528.23
    let code = r#"
        Sub AutoOpen()
            Dim result As Double
            result = FV(0.05/12, 120, -100, 0, 0)
            MsgBox Round(result, 2)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "15528.23");
}

#[test]
fn test_fv_zero_rate() {
    // No interest: just sum of payments
    // FV(0, 12, -100, 0, 0) = 1200
    let code = r#"
        Sub AutoOpen()
            MsgBox FV(0, 12, -100, 0, 0)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "1200");
}

#[test]
fn test_pv_monthly_investment() {
    // PV(Rate, NPer, Pmt, [FV], [Due])
    // $100/month at 5% annual for 10 years
    // PV(0.05/12, 120, -100, 0, 0) ≈ 9428.14
    let code = r#"
        Sub AutoOpen()
            Dim result As Double
            result = PV(0.05/12, 120, -100, 0, 0)
            MsgBox Round(result, 2)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "9428.14");
}

#[test]
fn test_npv_cash_flows() {
    // NPV(Rate, Value1, Value2, ...)
    // NPV(0.1, -1000, 200, 300, 400, 500) ≈ 65.26
    let code = r#"
        Sub AutoOpen()
            Dim result As Double
            result = NPV(0.1, -1000, 200, 300, 400, 500)
            MsgBox Round(result, 2)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "65.26");
}

// ============================================================
// PAYMENT FUNCTION TESTS
// ============================================================

#[test]
fn test_pmt_mortgage() {
    // Pmt(Rate, NPer, PV, [FV], [Due])
    // $200,000 mortgage at 6% for 30 years
    // Pmt(0.06/12, 360, 200000, 0, 0) ≈ -1199.10
    let code = r#"
        Sub AutoOpen()
            Dim result As Double
            result = Pmt(0.06/12, 360, 200000, 0, 0)
            MsgBox Round(result, 2)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "-1199.1");
}

#[test]
fn test_pmt_zero_rate() {
    // No interest: equal payments
    // Pmt(0, 12, 1200, 0, 0) = -100
    let code = r#"
        Sub AutoOpen()
            MsgBox Pmt(0, 12, 1200, 0, 0)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "-100");
}

#[test]
fn test_ipmt_first_period() {
    // IPmt(Rate, Per, NPer, PV, [FV], [Due])
    // Interest portion of first payment on $10000 at 10% for 5 years
    let code = r#"
        Sub AutoOpen()
            Dim result As Double
            result = IPmt(0.1/12, 1, 60, 10000, 0, 0)
            MsgBox Round(result, 2)
        End Sub
    "#;
    let result = run_vba_first(code);
    // First month interest on $10000 at 10%/12 ≈ $83.33
    assert!(result.contains("83.33"));
}

#[test]
fn test_ppmt_first_period() {
    // PPmt(Rate, Per, NPer, PV, [FV], [Due])
    // Principal portion of first payment
    let code = r#"
        Sub AutoOpen()
            Dim result As Double
            result = PPmt(0.1/12, 1, 60, 10000, 0, 0)
            MsgBox Round(result, 2)
        End Sub
    "#;
    let result = run_vba_first(code);
    // First month principal is a negative payment - the value will be around -129 to -300
    let val: f64 = result.parse().unwrap_or(0.0);
    assert!(val < 0.0, "PPMT should be negative for a loan");
    assert!(val > -400.0 && val < -50.0, "PPMT should be between -50 and -400");
}

#[test]
fn test_ipmt_plus_ppmt_equals_pmt() {
    // IPmt + PPmt should equal Pmt for any period
    let code = r#"
        Sub AutoOpen()
            Dim rate As Double
            Dim pmt_val As Double
            Dim ipmt_val As Double
            Dim ppmt_val As Double
            
            rate = 0.08 / 12
            pmt_val = Pmt(rate, 48, 20000, 0, 0)
            ipmt_val = IPmt(rate, 12, 48, 20000, 0, 0)
            ppmt_val = PPmt(rate, 12, 48, 20000, 0, 0)
            
            ' Check if IPmt + PPmt = Pmt
            MsgBox Round(Abs(pmt_val - (ipmt_val + ppmt_val)), 6)
        End Sub
    "#;
    let result = run_vba_first(code);
    // Should be 0 or very close to 0
    let val: f64 = result.parse().unwrap_or(999.0);
    assert!(val < 0.001, "IPmt + PPmt should equal Pmt");
}

// ============================================================
// LOAN/INVESTMENT FUNCTION TESTS
// ============================================================

#[test]
fn test_nper_loan() {
    // NPer(Rate, Pmt, PV, [FV], [Due])
    // How many months to pay off $50,000 at 5% with $500/month
    // NPer(0.05/12, -500, 50000, 0, 0) ≈ 129.63
    let code = r#"
        Sub AutoOpen()
            Dim result As Double
            result = NPer(0.05/12, -500, 50000, 0, 0)
            MsgBox Round(result, 2)
        End Sub
    "#;
    let result = run_vba_first(code);
    // Accept either 129.63 or 129.64 due to rounding
    assert!(result == "129.63" || result == "129.64");
}

#[test]
fn test_nper_zero_rate() {
    // Zero interest rate case
    // NPer(0, -100, 1200, 0, 0) = 12
    let code = r#"
        Sub AutoOpen()
            MsgBox NPer(0, -100, 1200, 0, 0)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "12");
}

#[test]
fn test_rate_loan() {
    // Rate(NPer, Pmt, PV, [FV], [Due], [Guess])
    // Find rate for 60 months, -$500/month, $25,000 loan
    let code = r#"
        Sub AutoOpen()
            Dim result As Double
            result = Rate(60, -500, 25000, 0, 0, 0.1)
            MsgBox Round(result, 5)
        End Sub
    "#;
    let result = run_vba_first(code);
    // Should be approximately 0.618% per month
    let val: f64 = result.parse().unwrap_or(0.0);
    assert!((val - 0.00618).abs() < 0.001, "Rate should be approximately 0.618%");
}

// ============================================================
// INTERNAL RATE OF RETURN FUNCTION TESTS
// ============================================================

#[test]
fn test_irr_investment() {
    // IRR(CashFlow1, CashFlow2, ...)
    // Initial investment -$10,000 with returns 3000, 4000, 4000, 3000
    // IRR should be approximately 14.89%
    let code = r#"
        Sub AutoOpen()
            Dim result As Double
            result = IRR(-10000, 3000, 4000, 4000, 3000)
            MsgBox Round(result, 4)
        End Sub
    "#;
    let result = run_vba_first(code);
    let val: f64 = result.parse().unwrap_or(0.0);
    assert!((val - 0.1489).abs() < 0.01, "IRR should be approximately 14.89%");
}

#[test]
fn test_irr_break_even() {
    // IRR where NPV = 0 at ~10%
    // -1000, 550, 605 (returns sum to 1155, should yield ~10% IRR)
    let code = r#"
        Sub AutoOpen()
            Dim result As Double
            result = IRR(-1000, 550, 605)
            MsgBox Round(result, 3)
        End Sub
    "#;
    let result = run_vba_first(code);
    let val: f64 = result.parse().unwrap_or(0.0);
    assert!(val > 0.09 && val < 0.12, "IRR should be close to 10%");
}

#[test]
fn test_mirr_investment() {
    // MIRR(CashFlow1, ..., FinanceRate, ReinvestRate)
    // -10000, 3000, 4000, 4000, 3000 with 10% finance and 12% reinvest
    let code = r#"
        Sub AutoOpen()
            Dim result As Double
            result = MIRR(-10000, 3000, 4000, 4000, 3000, 0.1, 0.12)
            MsgBox Round(result, 4)
        End Sub
    "#;
    let result = run_vba_first(code);
    let val: f64 = result.parse().unwrap_or(0.0);
    // MIRR should be between 10% and 15%
    assert!(val > 0.10 && val < 0.20, "MIRR should be a reasonable return rate");
}

// ============================================================
// EDGE CASE TESTS
// ============================================================

#[test]
fn test_pmt_due_at_beginning() {
    // Due = 1 means payments at beginning of period
    let code = r#"
        Sub AutoOpen()
            Dim pmt_end As Double
            Dim pmt_begin As Double
            
            pmt_end = Pmt(0.05/12, 60, 10000, 0, 0)
            pmt_begin = Pmt(0.05/12, 60, 10000, 0, 1)
            
            ' Payment at beginning should be slightly smaller (less interest)
            MsgBox pmt_begin > pmt_end
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "True");
}

#[test]
fn test_fv_with_initial_pv() {
    // FV with initial investment
    // Starting with $5000, adding $100/month for 5 years at 6%
    let code = r#"
        Sub AutoOpen()
            Dim result As Double
            result = FV(0.06/12, 60, -100, -5000, 0)
            MsgBox Round(result, 2)
        End Sub
    "#;
    let result = run_vba_first(code);
    // Should be greater than 5000 + 6000 = 11000 due to interest
    let val: f64 = result.parse().unwrap_or(0.0);
    assert!(val > 13000.0, "FV should include growth on initial investment");
}

#[test]
fn test_npv_single_cash_flow() {
    // NPV with single cash flow
    // NPV(0.1, 1100) = 1100 / 1.1 = 1000
    let code = r#"
        Sub AutoOpen()
            MsgBox Round(NPV(0.1, 1100), 0)
        End Sub
    "#;
    let result = run_vba_first(code);
    assert_eq!(result, "1000");
}

#[test]
fn test_financial_functions_integration() {
    // Full integration test: calculate mortgage details
    let code = r#"
        Sub AutoOpen()
            Dim principal As Double
            Dim annual_rate As Double
            Dim monthly_rate As Double
            Dim years As Integer
            Dim months As Integer
            Dim monthly_payment As Double
            
            principal = 250000
            annual_rate = 0.065
            monthly_rate = annual_rate / 12
            years = 30
            months = years * 12
            
            monthly_payment = -Pmt(monthly_rate, months, principal, 0, 0)
            
            ' Verify by calculating FV of payments equals 0
            Dim fv_check As Double
            fv_check = FV(monthly_rate, months, -monthly_payment, principal, 0)
            
            MsgBox Round(Abs(fv_check), 2)
        End Sub
    "#;
    let result = run_vba_first(code);
    // FV should be very close to 0 (loan paid off)
    let val: f64 = result.parse().unwrap_or(999.0);
    assert!(val < 0.01, "FV of mortgage payments should equal 0");
}
