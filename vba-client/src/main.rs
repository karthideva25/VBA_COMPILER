// use tree_sitter::Parser;
// use vba_parser::language;

// fn main() {
//     let mut parser = Parser::new();
//     parser.set_language(language()).unwrap();

//     let tree = parser.parse("Sub Hello()\nMsgBox \"Hi\"\nEnd Sub", None).unwrap();
//     println!("{}", tree.root_node().to_sexp());
// }

// vba-client/src/main.rs
use tree_sitter::Parser;
use vba_parser::language as tree_sitter_vba;
use vba_utils::{print_parse_tree, Context, Program, RuntimeConfig};
use vba_utils::ast::Statement;
use vba_utils::vm::ProgramExecutor; // ✅ import ProgramExecutor

fn main() {
    let vba_code = r#"

    Sub AutoOpen()
    ' Entry point - no parameters
    MsgBox "=== AutoOpen Started ==="
    Worksheets("Sheet1").Range("a1").Value = "hello from AutoOpen"
    
    ' Test Date/Time Functions
    Call TestDateTimeFunctions
    
    ' Test Missing Date/Time Functions
    Call TestMissingDateTimeFunctions
    
    ' Test Financial Functions
    Call TestFinancialFunctions
    
    MsgBox "=== AutoOpen Completed ==="
End Sub

' ============================================
' Test Date/Time Functions
' ============================================
Sub TestDateTimeFunctions()
    Dim d As Variant
    Dim dt As Variant
    Dim result As Variant
    
    MsgBox "=== Date/Time Function Tests ===" & vbCrLf
    MsgBox "Testing with user timezone from RuntimeConfig" & vbCrLf
    
    ' ============================================
    ' Current Date/Time Functions
    ' ============================================
    MsgBox "--- Current Date/Time ---"
    MsgBox "Now() = " & Now()
    MsgBox "Date() = " & Date()
    MsgBox "Time() = " & Time()
    
    ' ============================================
    ' Time Arithmetic Functions
    ' ============================================
    MsgBox "--- Time Arithmetic ---"
    MsgBox "TimeValue(""00:00:10"") = " & TimeValue("00:00:10")
    MsgBox "TimeValue(""01:30:00"") = " & TimeValue("01:30:00")
    MsgBox "TimeSerial(14, 30, 45) = " & TimeSerial(14, 30, 45)
    
    ' Adding time to Now()
    MsgBox "Now() + TimeValue(""00:00:10"") = " & (Now() + TimeValue("00:00:10"))
    MsgBox "Now() + TimeValue(""01:00:00"") = " & (Now() + TimeValue("01:00:00"))
    MsgBox "Now() + TimeValue(""00:30:00"") = " & (Now() + TimeValue("00:30:00"))
    
    ' Adding time to a Date
    d = #2/15/2026#
    MsgBox "Date + Time: " & d & " + TimeValue(""10:30:00"") = " & (d + TimeValue("10:30:00"))
    
    ' Time + Time
    MsgBox "TimeValue(""01:30:00"") + TimeValue(""00:45:30"") = " & (TimeValue("01:30:00") + TimeValue("00:45:30"))
    
    ' ============================================
    ' Time Part Extraction Functions
    ' ============================================
    MsgBox "--- Time Part Extraction from Now() ---"
    MsgBox "Hour(Now()) = " & Hour(Now())
    MsgBox "Minute(Now()) = " & Minute(Now())
    MsgBox "Second(Now()) = " & Second(Now())
    
    MsgBox "--- Time Part Extraction from Time() ---"
    MsgBox "Hour(Time()) = " & Hour(Time())
    MsgBox "Minute(Time()) = " & Minute(Time())
    MsgBox "Second(Time()) = " & Second(Time())
    
    ' ============================================
    ' Date Part Extraction Functions
    ' ============================================
    MsgBox "--- Date Part Extraction ---"
    d = #2/15/2026#
    MsgBox "Test date: " & d
    MsgBox "Year(d) = " & Year(d)
    MsgBox "Month(d) = " & Month(d)
    MsgBox "Day(d) = " & Day(d)
    MsgBox "Weekday(d) = " & Weekday(d)
    
    ' Test with Now()
    MsgBox "--- Parts from Now() ---"
    MsgBox "Year(Now()) = " & Year(Now())
    MsgBox "Month(Now()) = " & Month(Now())
    MsgBox "Day(Now()) = " & Day(Now())
    MsgBox "Weekday(Now()) = " & Weekday(Now())
    
    ' ============================================
    ' Date Construction Functions
    ' ============================================
    MsgBox "--- Date Construction ---"
    result = DateSerial(2026, 7, 4)
    MsgBox "DateSerial(2026, 7, 4) = " & result
    
    result = DateSerial(2026, 1, 1)
    MsgBox "DateSerial(2026, 1, 1) = " & result
    
    result = DateSerial(2026, 12, 31)
    MsgBox "DateSerial(2026, 12, 31) = " & result
    
    ' Edge cases - month/day rollover
    result = DateSerial(2026, 13, 1)
    MsgBox "DateSerial(2026, 13, 1) = " & result & " (month overflow)"
    
    result = DateSerial(2026, 2, 30)
    MsgBox "DateSerial(2026, 2, 30) = " & result & " (day overflow)"
    
    ' ============================================
    ' DateValue - Parse date from string
    ' ============================================
    MsgBox "--- DateValue (String to Date) ---"
    result = DateValue("2026-03-15")
    MsgBox "DateValue(""2026-03-15"") = " & result
    
    result = DateValue("03/15/2026")
    MsgBox "DateValue(""03/15/2026"") = " & result
    
    ' ============================================
    ' WeekdayName and MonthName
    ' ============================================
    MsgBox "--- WeekdayName and MonthName ---"
    MsgBox "WeekdayName(1, False) = " & WeekdayName(1, False)
    MsgBox "WeekdayName(2, False) = " & WeekdayName(2, False)
    MsgBox "WeekdayName(3, True) = " & WeekdayName(3, True)
    MsgBox "WeekdayName(4, True) = " & WeekdayName(4, True)
    MsgBox "WeekdayName(5, False) = " & WeekdayName(5, False)
    MsgBox "WeekdayName(6, False) = " & WeekdayName(6, False)
    MsgBox "WeekdayName(7, False) = " & WeekdayName(7, False)
    
    MsgBox "MonthName(1, False) = " & MonthName(1, False)
    MsgBox "MonthName(6, False) = " & MonthName(6, False)
    MsgBox "MonthName(12, False) = " & MonthName(12, False)
    MsgBox "MonthName(3, True) = " & MonthName(3, True)
    
    ' ============================================
    ' DateAdd - Add intervals to dates
    ' ============================================
    MsgBox "--- DateAdd ---"
    d = #1/15/2026#
    MsgBox "Base date: " & d
    
    result = DateAdd("d", 10, d)
    MsgBox "DateAdd(""d"", 10, d) = " & result & " (add 10 days)"
    
    result = DateAdd("m", 3, d)
    MsgBox "DateAdd(""m"", 3, d) = " & result & " (add 3 months)"
    
    result = DateAdd("yyyy", 1, d)
    MsgBox "DateAdd(""yyyy"", 1, d) = " & result & " (add 1 year)"
    
    result = DateAdd("ww", 2, d)
    MsgBox "DateAdd(""ww"", 2, d) = " & result & " (add 2 weeks)"
    
    result = DateAdd("d", -5, d)
    MsgBox "DateAdd(""d"", -5, d) = " & result & " (subtract 5 days)"
    
    ' ============================================
    ' DateDiff - Difference between dates
    ' ============================================
    MsgBox "--- DateDiff ---"
    Dim d1 As Variant
    Dim d2 As Variant
    d1 = #1/1/2026#
    d2 = #3/15/2026#
    MsgBox "Date1: " & d1 & ", Date2: " & d2
    
    result = DateDiff("d", d1, d2)
    MsgBox "DateDiff(""d"", d1, d2) = " & result & " days"
    
    result = DateDiff("m", d1, d2)
    MsgBox "DateDiff(""m"", d1, d2) = " & result & " months"
    
    result = DateDiff("ww", d1, d2)
    MsgBox "DateDiff(""ww"", d1, d2) = " & result & " weeks"
    
    d2 = #1/1/2027#
    result = DateDiff("yyyy", d1, d2)
    MsgBox "DateDiff(""yyyy"", d1, " & d2 & ") = " & result & " years"
    
    ' ============================================
    ' DatePart - Extract specific part of date
    ' ============================================
    MsgBox "--- DatePart ---"
    d = #7/4/2026#
    MsgBox "Test date: " & d
    
    result = DatePart("yyyy", d)
    MsgBox "DatePart(""yyyy"", d) = " & result
    
    result = DatePart("m", d)
    MsgBox "DatePart(""m"", d) = " & result
    
    result = DatePart("d", d)
    MsgBox "DatePart(""d"", d) = " & result
    
    result = DatePart("w", d)
    MsgBox "DatePart(""w"", d) = " & result & " (weekday)"
    
    result = DatePart("ww", d)
    MsgBox "DatePart(""ww"", d) = " & result & " (week of year)"
    
    result = DatePart("q", d)
    MsgBox "DatePart(""q"", d) = " & result & " (quarter)"
    
    result = DatePart("y", d)
    MsgBox "DatePart(""y"", d) = " & result & " (day of year)"
    
    ' ============================================
    ' FormatDateTime - Format date with named formats
    ' ============================================
    MsgBox "--- FormatDateTime ---"
    d = #11/25/2026#
    MsgBox "Test date: " & d
    
    result = FormatDateTime(d, 0)
    MsgBox "FormatDateTime(d, 0) = " & result & " (General)"
    
    result = FormatDateTime(d, 1)
    MsgBox "FormatDateTime(d, 1) = " & result & " (Long Date)"
    
    result = FormatDateTime(d, 2)
    MsgBox "FormatDateTime(d, 2) = " & result & " (Short Date)"
    
    result = FormatDateTime(d, 3)
    MsgBox "FormatDateTime(d, 3) = " & result & " (Long Time)"
    
    result = FormatDateTime(d, 4)
    MsgBox "FormatDateTime(d, 4) = " & result & " (Short Time)"
    
    ' ============================================
    ' Format with custom date patterns
    ' ============================================
    MsgBox "--- Format with Custom Patterns ---"
    d = Now()
    
    result = Format(d, "yyyy-mm-dd")
    MsgBox "Format(d, ""yyyy-mm-dd"") = " & result
    
    result = Format(d, "dd/mm/yyyy")
    MsgBox "Format(d, ""dd/mm/yyyy"") = " & result
    
    result = Format(d, "yyyy/mm/dd HH:MM:SS")
    MsgBox "Format(d, ""yyyy/mm/dd HH:MM:SS"") = " & result
    
    result = Format(d, "mmmm dd, yyyy")
    MsgBox "Format(d, ""mmmm dd, yyyy"") = " & result
    
    result = Format(d, "ddd, mmm d")
    MsgBox "Format(d, ""ddd, mmm d"") = " & result
    
    result = Format(d, "Long Date")
    MsgBox "Format(d, ""Long Date"") = " & result
    
    result = Format(d, "Short Date")
    MsgBox "Format(d, ""Short Date"") = " & result
    
    ' ============================================
    ' IsDate - Check if value is a valid date
    ' ============================================
    MsgBox "--- IsDate ---"
    MsgBox "IsDate(#1/15/2026#) = " & IsDate(#1/15/2026#)
    MsgBox "IsDate(Now()) = " & IsDate(Now())
    MsgBox "IsDate(""2026-01-15"") = " & IsDate("2026-01-15")
    MsgBox "IsDate(""not a date"") = " & IsDate("not a date")
    MsgBox "IsDate(12345) = " & IsDate(12345)
    MsgBox "IsDate("""") = " & IsDate("")
    
    ' ============================================
    ' Date Calculations
    ' ============================================
    MsgBox "--- Date Calculations ---"
    
    ' Days until end of year
    Dim today As Variant
    Dim endOfYear As Variant
    today = Now()
    endOfYear = DateSerial(Year(today), 12, 31)
    result = DateDiff("d", today, endOfYear)
    MsgBox "Days until end of " & Year(today) & ": " & result
    
    ' What day is 100 days from now?
    result = DateAdd("d", 100, today)
    MsgBox "100 days from today: " & result
    MsgBox "That will be a " & WeekdayName(Weekday(result), False)
    
    ' First Monday of next month
    Dim nextMonth As Variant
    nextMonth = DateSerial(Year(today), Month(today) + 1, 1)
    MsgBox "First of next month: " & nextMonth & " (" & WeekdayName(Weekday(nextMonth), False) & ")"
    
    MsgBox "=== Date/Time Tests Complete ===" & vbCrLf
End Sub

' ============================================
' Test Missing Date/Time Functions
' ============================================
Sub TestMissingDateTimeFunctions()
    MsgBox "=== Missing DateTime Function Tests ===" & vbCrLf
    
    ' ===== TIMER =====
    MsgBox "--- Timer ---"
    MsgBox "Timer() = " & Timer()
    
    ' ===== NEGATIVE DATESERIAL =====
    MsgBox "--- Negative DateSerial ---"
    MsgBox "DateSerial(2026, -1, 1) = " & DateSerial(2026, -1, 1)
    MsgBox "DateSerial(2026, 0, 1) = " & DateSerial(2026, 0, 1)
    MsgBox "DateSerial(2026, 1, -5) = " & DateSerial(2026, 1, -5)
    
    ' ===== TIMESERIAL OVERFLOW =====
    MsgBox "--- TimeSerial Overflow ---"
    MsgBox "TimeSerial(25, 0, 0) = " & TimeSerial(25, 0, 0)
    MsgBox "TimeSerial(-1, 30, 0) = " & TimeSerial(-1, 30, 0)
    MsgBox "TimeSerial(0, 90, 0) = " & TimeSerial(0, 90, 0)
    MsgBox "TimeSerial(0, 0, 7200) = " & TimeSerial(0, 0, 7200)
    
    ' ===== DATEADD WITH TIME =====
    MsgBox "--- DateAdd with Time ---"
    Dim d As Variant
    ' Note: Parser doesn't support datetime literals with time, using DateSerial + TimeSerial
    d = DateSerial(2026, 1, 15) + TimeSerial(10, 30, 0)
    MsgBox "Base datetime: " & d
    MsgBox "DateAdd(""h"", 5, d) = " & DateAdd("h", 5, d)
    MsgBox "DateAdd(""n"", 30, d) = " & DateAdd("n", 30, d)
    MsgBox "DateAdd(""s"", 45, d) = " & DateAdd("s", 45, d)
    
    ' ===== DATEDIFF WITH TIME =====
    MsgBox "--- DateDiff with Time ---"
    Dim d1 As Variant, d2 As Variant
    ' Create datetimes using DateSerial + TimeSerial
    d1 = DateSerial(2026, 1, 1) + TimeSerial(10, 0, 0)
    d2 = DateSerial(2026, 1, 1) + TimeSerial(15, 30, 45)
    MsgBox "d1 = " & d1 & ", d2 = " & d2
    MsgBox "DateDiff(""h"", d1, d2) = " & DateDiff("h", d1, d2)
    MsgBox "DateDiff(""n"", d1, d2) = " & DateDiff("n", d1, d2)
    MsgBox "DateDiff(""s"", d1, d2) = " & DateDiff("s", d1, d2)
    
    ' ===== LEAP YEAR =====
    MsgBox "--- Leap Year Tests ---"
    MsgBox "DateSerial(2024, 2, 29) = " & DateSerial(2024, 2, 29)
    MsgBox "DateSerial(2025, 2, 29) = " & DateSerial(2025, 2, 29)
    
    ' ===== FORMAT TIME =====
    MsgBox "--- Format Time ---"
    Dim dt As Variant
    ' Create datetime using DateSerial + TimeSerial
    dt = DateSerial(2026, 2, 2) + TimeSerial(14, 5, 9)
    MsgBox "Test datetime: " & dt
    MsgBox "Format(dt, ""hh:nn:ss"") = " & Format(dt, "hh:nn:ss")
    MsgBox "Format(dt, ""h:n:s AM/PM"") = " & Format(dt, "h:n:s AM/PM")
    
    MsgBox "=== Missing DateTime Tests Complete ===" & vbCrLf
End Sub

' ============================================
' Test Financial Functions
' ============================================
Sub TestFinancialFunctions()
    MsgBox "=== Financial Function Tests ===" & vbCrLf
    
    ' ============================================
    ' DEPRECIATION FUNCTIONS
    ' ============================================
    MsgBox "--- Depreciation Functions ---"
    
    ' SLN - Straight Line Depreciation
    ' Asset cost $10,000, salvage $1,000, life 5 years
    MsgBox "SLN(10000, 1000, 5) = " & SLN(10000, 1000, 5)
    
    ' SYD - Sum of Years Digits
    MsgBox "SYD(10000, 1000, 5, 1) = " & SYD(10000, 1000, 5, 1)
    MsgBox "SYD(10000, 1000, 5, 3) = " & SYD(10000, 1000, 5, 3)
    
    ' DDB - Double Declining Balance
    MsgBox "DDB(10000, 1000, 5, 1) = " & DDB(10000, 1000, 5, 1)
    MsgBox "DDB(10000, 1000, 5, 2) = " & DDB(10000, 1000, 5, 2)
    
    ' ============================================
    ' PRESENT/FUTURE VALUE FUNCTIONS
    ' ============================================
    MsgBox "--- Present/Future Value Functions ---"
    
    ' FV - Future Value
    ' 5% annual rate, 10 years, $100/month payment
    MsgBox "FV(0.05/12, 120, -100, 0, 0) = " & FV(0.05/12, 120, -100, 0, 0)
    
    ' PV - Present Value  
    ' 5% rate, 10 years, $100/month, no FV
    MsgBox "PV(0.05/12, 120, -100, 0, 0) = " & PV(0.05/12, 120, -100, 0, 0)
    
    ' NPV - Net Present Value
    ' 10% discount rate, cash flows: -1000, 200, 300, 400, 500
    MsgBox "NPV(0.1, -1000, 200, 300, 400, 500) = " & NPV(0.1, -1000, 200, 300, 400, 500)
    
    ' ============================================
    ' PAYMENT FUNCTIONS
    ' ============================================
    MsgBox "--- Payment Functions ---"
    
    ' PMT - Payment for loan
    ' 6% annual rate, 30 year mortgage, $200,000 loan
    MsgBox "Pmt(0.06/12, 360, 200000, 0, 0) = " & Pmt(0.06/12, 360, 200000, 0, 0)
    
    ' IPMT - Interest portion of payment
    MsgBox "IPmt(0.06/12, 1, 360, 200000, 0, 0) = " & IPmt(0.06/12, 1, 360, 200000, 0, 0)
    MsgBox "IPmt(0.06/12, 12, 360, 200000, 0, 0) = " & IPmt(0.06/12, 12, 360, 200000, 0, 0)
    
    ' PPMT - Principal portion of payment
    MsgBox "PPmt(0.06/12, 1, 360, 200000, 0, 0) = " & PPmt(0.06/12, 1, 360, 200000, 0, 0)
    MsgBox "PPmt(0.06/12, 12, 360, 200000, 0, 0) = " & PPmt(0.06/12, 12, 360, 200000, 0, 0)
    
    ' ============================================
    ' LOAN/INVESTMENT FUNCTIONS
    ' ============================================
    MsgBox "--- Loan/Investment Functions ---"
    
    ' NPER - Number of periods
    ' 5% rate, $500/month, $50,000 loan
    MsgBox "NPer(0.05/12, -500, 50000, 0, 0) = " & NPer(0.05/12, -500, 50000, 0, 0)
    
    ' RATE - Interest rate per period
    ' 60 payments, $500/month, $25,000 loan
    MsgBox "Rate(60, -500, 25000, 0, 0, 0.1) = " & Rate(60, -500, 25000, 0, 0, 0.1)
    
    ' ============================================
    ' INTERNAL RATE OF RETURN
    ' ============================================
    MsgBox "--- Internal Rate of Return ---"
    
    ' IRR - Internal Rate of Return
    ' Initial investment -10000, returns: 3000, 4000, 4000, 3000
    MsgBox "IRR(-10000, 3000, 4000, 4000, 3000) = " & IRR(-10000, 3000, 4000, 4000, 3000)
    
    MsgBox "=== Financial Tests Complete ===" & vbCrLf
End Sub

    "#;

    // Set up parser
    let mut parser = Parser::new();
    parser
        .set_language(tree_sitter_vba())
        .expect("Failed to load grammar");

    // Parse input
    let tree = parser.parse(vba_code, None).expect("Parsing failed");
    let root_node = tree.root_node();

    // println!("\n🔍 Tree-sitter Parse Tree:");
   // print_parse_tree(vba_code); // <-- You already have a utility for this!

    // Build AST from the parse tree
    use vba_utils::ast::build_ast;
    let program: Program = build_ast(root_node, vba_code);

    // DUMP THE WHOLE AST
    // dbg!(&program);

    // ============================================================
    // Create Context with RuntimeConfig
    // In production: get these values from user session/profile
    // ============================================================
    let config = RuntimeConfig::builder()
        .timezone("Asia/Kolkata")       // User's timezone from profile
        .locale("en-IN")                // User's locale
        .workbook_id("sample-workbook") // Active workbook ID
        .user_id("user-12345")          // Authenticated user
        .build();
    
    let mut ctx = Context::with_config(config);
    println!("\n🧠 Interpreting AST (Timezone: {}):", ctx.runtime_config.timezone_name());

    // ✅ Use the new 3-phase executor
    let executor = ProgramExecutor::new(program);
    if let Err(e) = executor.execute(&mut ctx) {
        eprintln!("Program execution error: {}", e);
    }
    
    // Print workbook ID and cell value [0,0] (A1) from active workbook
    println!("\n📊 Checking workbook and cell value:");
    
    // Get workbook ID
    let workbook_id = vba_utils::host::excel::engine::get_workbook_id();
    match workbook_id {
        Some(rid) => println!("🆔 Workbook ID (rid): {}", rid),
        None => println!("⚠️  No active workbook"),
    }
    
    // Get cell value [0,0]
    match vba_utils::host::excel::engine::get_cell_value("A1") {
        Ok(value) => println!("✅ Cell A1 value: {}", value),
        Err(e) => println!("❌ Error reading A1: {}", e),
    }
}
//Sub somemacro()

// Dim j As Integer
// Dim h As Integer
// Dim a as Range
// set a=Range("a1")
// j = 10
// h = 45
// variable = "variable"
// variable = variable + j
// MsgBox (variable )
// Range("A1").Value
// call Integer
// call Range
// If j<100 Then
//     variable = "ten"
// Else 
//     variable = "not ten"
// End If

// End Sub


//     Sub HelloWorld()
//         MsgBox "Hello, World!"
//         Dim j As Integer
//         Dim h As Integer
//         j = 10
//         h = 45
//         If j Then
//             variable = "ten"
//         Else If
//             variable = "not ten"
//         End If

//     End Sub

// Sub B()
// A "john"
// End Sub
// Sub A(name)
// MsgBox "Hello, World!"
// Dim const as Variable
// Dim b as Const 
// Dim a as Variant
// Dim h as Integer
// End Sub

// Range("A1").Value = Cell
// MsgBox ("variable" )
// If j<100 Then
//     MsgBox "ten"
// Else 
//     MsgBox "not ten"
// End If


// Sub somemacro()

//         Dim j As Integer
//         Dim h As Integer
//         Dim m As Integer
//         m= 79
//         j = 10+ m +1000
//         h = vbRed
//         If j<100 Then
//              MsgBox "ten" & j
//         Else 
//              MsgBox "not ten"
//         End If
//         Dim i As Integer
//         i=1
//         For i To 5
//             MsgBox "Value of i = " & i
//         Next i
//         j=78
//         MsgBox h
//         Msgbox " I am a msgbox msg &j"
//         Msgbox " I am a msgbox msg" & j
//         MsgBox ("red " & vbRed)
       
//     End Sub

// Sub somemacro()
// Dim i As Integer
// For i = 1 To 5
// If i = 3 Then
//     Msgbox "Inside if" & i
//     Msgbox "Inside if  i should be 3 : " & i
    
//     Exit For
//     Msgbox "Inside if for exited" & i
// End If 
// Msgbox "count first " &i
// Next i
// For i = 1 To 5
// If i = 3 Then
//     GoTo AfterFor
// End If
// MsgBox "count" & i
// Next i
// AfterFor:
// MsgBox "done"


// MsgBox "I am some" & vbCrLf & "vbCrLf"

// End Sub


// Sub AutoOpen()  
//         Call TestDoWhileGoTo
//         'Call UseEmployeeType
//         MsgBox  " i am  in auto open "
//         Call TestDoWhileWithErrorHandler
//      End Sub
    
//     Sub TestDoWhileGoTo()
//         Dim i As Integer
//         i = 1
        
//         Do While i <= 5
//             MsgBox i
//             If i = 7 Then
//                 GoTo ExitPoint
//             End If
//             i = i + 1
//         Loop
        
//         MsgBox "Should not reach here"
        
//     ExitPoint:
//         MsgBox "Jumped out"
//     End Sub

//     Sub TestDoWhileWithErrorHandler()

//         On Error GoTo ErrHandler   ' Enable error handling

//         Dim i As Integer
//         i = 1

//         Do While i <= 5
//             MsgBox "Loop iteration: " & i

//             ' Intentional test error (divide by zero when i = 3)
//             If i = 3 Then
//                 MsgBox " error" + 10 / 0
//             End If

//             i = i + 1
//         Loop

//         Exit Sub    ' Prevents running into the handler when no error occurs

//     ErrHandler:
//         MsgBox "Error occurred: " & Err.Number & " - " & Err.Description
//         Resume Next     ' Continue with the next line after the one that caused the error

//     End Sub
