// Tree-sitter grammar for VBA - Enhanced Assignment Statement

// Helper functions for comma-separated lists
function commaSep(rule) {
  return sep1(rule, ',');
}
function commaSep1(rule) {
  return seq(rule, repeat(seq(',', rule)));
}

function sep1(rule, separator) {
  return seq(rule, repeat(seq(separator, rule)));
}

const PREC = {
  call: 16,        // function call
  member: 15,      // property access, array indexing
  mul: 13,         // *, /
  add: 12,         // +, -
  concat: 11,      // &
  relational: 10,  // >, <, >=, <=
  equality: 9,     // =, <>
  assignment: 1    // assignment has low precedence
};

module.exports = grammar({
  name: 'vba',

  // Declare conflict sets to resolve ambiguous newline consumption
  conflicts: $ => [
    [$.call_statement, $.blank_line],
    [$.call_statement, $.expression_statement],
    [$.call_statement, $.expression],
    [$.subroutine, $.blank_line],
    [$.if_statement, $.blank_line],
    [$.if_statement],
    [$.lvalue, $.expression],  // Added for assignment targets
    [$.indexed_access, $.function_call],  // Added for array vs function disambiguation
    [$.qualified_identifier,$.property_access],
    [$.exit_statement, $.blank_line],
    [$.resume_statement, $.on_error_statement, $.for_statement],
  ],

  extras: $ => [
    /[ \t]+/,  // spaces and tabs
    /_/,        // line continuation
    $.comment,   // VBA comments
  ],

  rules: {
    source_file: $ => repeat($.statement),

    // A statement can be a blank line or any VBA statement
    statement: $ => choice(
      $.blank_line,
      $.option_explicit_statement,
      $.subroutine,
      $.dim_statement,
      $.redim_statement,
      $.enum_statement,  
      $.type_statement,
      $.set_statement,
      $.assignment_statement,
      $.msgbox_statement,
      $.goto_statement,
      $.if_statement,
      $.for_statement,
      $.do_while_statement,
      $.label_statement,
      $.expression_statement,
      $.call_statement,
      $.on_error_statement,
      $.resume_statement,
      $.exit_statement,

    ),

    // Blank line consumes a standalone newline
    blank_line: $ => /\r?\n/,

    // option_explicit_statement rule 
    option_explicit_statement: $ => seq(
      token(/Option/i),
      token(/Explicit/i),
      /\r?\n/
    ),

    // Subroutine Definition: Sub Name(params) ... End Sub
    subroutine: $ => seq(
      token(/Sub/i),
      field("name", $.identifier),
      optional($.parameter_list),
      /\r?\n/,
      repeat($.statement),
      token(/End\s+Sub/i),
      /\r?\n/
    ),

    // Parameter list: parentheses with comma-separated identifiers
    parameter_list: $ => seq(
      '(',
      optional(commaSep($.identifier)),
      ')'
    ),

    // Dim statement: Dim var [As Type]
    dim_statement: $ => seq(
      token(/Dim/i),
      commaSep(
        seq(
          field('name', $.identifier),
          optional(seq(choice(
            token(/as/i),token(/As/i)),
            field('type', choice(
                $.primitive_type,    // Byte, Integer, String, etc.
                $.identifier         // User-defined types like Employee
            ))
          ))
        )
      ),
      /\r?\n/
    ),

    redim_statement: $ => seq(
      token(/ReDim/i),
      optional(token(/Preserve/i)),
      commaSep(
        seq(
          field('name', $.identifier),
          '(',
          field('dimensions', commaSep($.array_dimension)),
          ')'
        )
      ),
      /\r?\n/
    ),
    // Add the enum_statement rule:
    enum_statement: $ => seq(
      // Optional visibility modifier (Public or Private)
      optional(field('visibility', choice(
        token(/Public/i),
        token(/Private/i)
      ))),
      token(/Enum/i),
      field('name', $.identifier),
      /\r?\n/,
      // One or more enum members
      repeat1($.enum_member),
      token(/End/i),
      token(/Enum/i),
      /\r?\n/
    ),

    // Enum member: membername [= constantexpression]
    enum_member: $ => seq(
      field('name', $.identifier),
      // Optional constant value assignment
      optional(seq(
        '=',
        field('value', $.expression)
      )),
      /\r?\n/
    ),
    // Set statement: Set var = expression
    set_statement: $ => seq(
      token(/Set/i),
      field('target', $.lvalue),  // Enhanced to use lvalue
      '=',
      field('value', $.expression),
      /\r?\n/
    ),
    // Add the type_statement rule:
    type_statement: $ => seq(
      // Optional visibility modifier (Public or Private)
      optional(field('visibility', choice(
        token(/Public/i),
        token(/Private/i)
      ))),
      token(/Type/i),
      field('name', $.identifier),
      /\r?\n/,
      // One or more type members (fields)
      repeat1($.type_member),
      token(/End/i),
      token(/Type/i),
      /\r?\n/
    ),

    // Type member: elementname [(subscripts)] As type
    type_member: $ => seq(
      field('name', $.identifier),
      // Optional array dimensions
      optional(seq(
        '(',
        field('dimensions', optional(commaSep($.array_dimension))),
        ')'
      )),
      token(/As/i),
      field('type', choice(
        $.primitive_type,
        $.identifier  // For user-defined types or Object
      )),
      optional(seq(
        '*',
        field('string_length', $.integer_literal)  // For fixed-length strings
      )),
      /\r?\n/
    ),
    // Array dimension: [lower To] upper or just index
    array_dimension: $ => choice(
      // Range: 1 To 100
      seq(
        field('lower', $.expression),
        token(/To/i),
        field('upper', $.expression)
      ),
      // Single value (implies 0 To value or just size)
      field('size', $.expression)
    ),

    // Enhanced Assignment: lvalue = expression
    // Supports: x = 5, obj.prop = value, arr(1) = data, obj.method().prop = val
    assignment_statement: $ => prec.right(PREC.assignment, seq(
      field('target', $.lvalue),
      optional(/[ \t]+/),
      token('='),    
      optional(/[ \t]+/),
      field('value', $.expression),
      /\r?\n/
    )),

    // Left-hand side values (what can be assigned to)
    lvalue: $ => choice(
      $.identifier,                    // Simple variable: x
      $.property_access,               // Object property: obj.prop
      $.indexed_access,                // Array/collection: arr(1), dict("key")
      $.qualified_identifier           // Module.variable (added)
    ),

    // Qualified identifier for module-level variables: Module.Variable
    qualified_identifier: $ => prec.left(PREC.member, seq(
      field('module', $.identifier),
      '.',
      field('name', $.identifier)
    )),

    // Enhanced indexed access for arrays and collections: obj(index1, index2, ...)
    indexed_access: $ => prec.left(PREC.member, seq(
      field('object', choice(
        $.identifier,
        $.property_access,
        $.indexed_access,
        $.function_call
      )),
      '(',
      field('indices', commaSep($.expression)),
      ')'
    )),

    // MsgBox statement: MsgBox expr
    msgbox_statement: $ => seq(
      token(/MsgBox/i),
      $.expression,
      /\r?\n/
    ),

    // GoTo statement: GoTo label
    goto_statement: $ => seq(
      token(/GoTo/i),
      $.identifier,
    ),

    if_statement: $ => choice(
      // Block form: newline after Then, requires End If
      seq(
        $.keyword_If,
        $.expression,
        $.keyword_Then,
        /\r?\n/,  // This newline makes it a BLOCK form
        
        repeat($.statement),
        
        repeat(seq(
          $.keyword_ElseIf,
          $.expression,
          $.keyword_Then,
          /\r?\n/,
          repeat($.statement)
        )),
        
        optional(seq(
          $.keyword_Else,
          /\r?\n/,
          repeat($.statement)
        )),
        
        // End If is REQUIRED for block form
        $.keyword_End_If,
        optional(/\r?\n/)
      ),
      
      // Inline form: NO newline after Then, NO End If
      prec.right(seq(
        $.keyword_If,
        $.expression,
        $.keyword_Then,
        // NO newline - stays on same line
        $.statement,
        
        optional(seq(
          $.keyword_Else,
          $.statement
        ))
        // NO End If allowed here
      ))
    ),   
    // For...Next statement: For counter = start To end [Step step]
    for_statement: $ => seq(
      token(/For/i),
      field('counter', $.identifier),
      '=',
      field('start', $.expression),
      token(/To/i),
      field('end', $.expression),
      // Optional Step clause
      optional(seq(
        token(/Step/i),
        field('step', $.expression)
      )),
      /\r?\n/,
      // Loop body - can contain any statements
      field('body', repeat($.statement)),
      token(/Next/i),
      // Optional counter variable after Next (can be omitted in VBA)
      optional(field('next_counter', $.identifier)),
      /\r?\n/
    ),

    do_while_statement: $ => choice(
      // Do While...Loop (condition at start)
      seq(
        token(/Do/i),
        token(/While/i),
        field('condition', $.expression),
        /\r?\n/,
        field('body', repeat($.statement)),
        token(/Loop/i),
        /\r?\n/
      ),
      
      // Do Until...Loop (condition at start)
      seq(
        token(/Do/i),
        token(/Until/i),
        field('condition', $.expression),
        /\r?\n/,
        field('body', repeat($.statement)),
        token(/Loop/i),
        /\r?\n/
      ),
      
      // Do...Loop While (condition at end)
      seq(
        token(/Do/i),
        /\r?\n/,
        field('body', repeat($.statement)),
        token(/Loop/i),
        token(/While/i),
        field('condition', $.expression),
        /\r?\n/
      ),
      
      // Do...Loop Until (condition at end)
      seq(
        token(/Do/i),
        /\r?\n/,
        field('body', repeat($.statement)),
        token(/Loop/i),
        token(/Until/i),
        field('condition', $.expression),
        /\r?\n/
      ),
      
      // Do...Loop (infinite loop, exit with Exit Do)
      seq(
        token(/Do/i),
        /\r?\n/,
        field('body', repeat($.statement)),
        token(/Loop/i),
        /\r?\n/
      )
    ),
        
   // Call statement: Call Func(args)
    call_statement: $ => seq(
      optional(token(/Call/i)),           // allow `Call Foo()` or just `Foo()`
      field("function", $.identifier),
      optional(choice(
        $.argument_list,             // e.g. Foo(a, b)
        seq(
          " ",                        // a space
          $.expression                // then a single expression, e.g. Foo "bar"
        )
      )),
      /\r?\n/                             // require statement-terminating newline
    ),

    // Label: Identifier:
    label_statement: $ => seq(
      $.identifier,
      ':',
      /\r?\n/
    ),

    // Expression-only statement
    expression_statement: $ => seq(
      prec(PREC.call - 1, $.expression),
      /\r?\n/
    ),

    // Argument list for calls: (expr, expr, ...)
    argument_list: $ => seq(
      '(',
      optional(commaSep($.expression)),
      ')'
    ),
    // Optional: more robust terminator handling
    _statement_terminator: $ => choice(/\r?\n/, token.immediate(':')),

    exit_statement: $ => seq(
      token(/Exit/i),
      field('exit_type', choice(
        token.immediate(/For/i),
        token.immediate(/Do/i),
        token.immediate(/While/i),
        token.immediate(/Sub/i),
        token.immediate(/Function/i),
        token.immediate(/Property/i),
        token.immediate(/Select/i)
      )),
      $._statement_terminator
    ),
    on_error_statement: $ => prec.left(seq(
      token(/On/i), token(/Error/i),
      choice(
        seq(token(/Resume/i), token(/Next/i)),
        seq(token(/GoTo/i), field('target', $.identifier)),
        seq(token(/GoTo/i), field('target', token.immediate('0')))
      ),
      optional($._statement_terminator)
    )),
    resume_statement: $ => prec.left(seq(
      token(/Resume/i),
      optional(choice(
        token(/Next/i),
        field('label', $.identifier)
      )),
      optional($._statement_terminator)
    )),

    // Enhanced expressions: includes all your existing ones plus object creation
    expression: $ => choice(
      $.binary_expression,
      $.unary_expression,
      $.function_call,
      $.property_access,
      $.indexed_access,        // Added: arrays can be expressions too
      $.object_creation,       // Added: New ClassName
      $.parenthesized_expression, // Added: (expr)
      $.vba_builtin_constant,  // Added: VBA built-in constants
      $.byte_literal,  
      $.integer_literal,
      $.string_literal,
      $.boolean_literal,       // Added: True/False
      $.date_literal,
      $.currency_literal,
      $.float_literal,
      $.nothing_literal,       // Added: Nothing
      $.identifier
    ),
    vba_builtin_constant: $ => choice(
      'vbCalGreg',
      'vbCalHijri',

      'vbMethod',
      'vbGet',
      'vbLet',
      'vbSet',

      // Color constants
      'vbBlack',
      'vbRed',
      'vbGreen',
      'vbYellow',
      'vbBlue',
      'vbMagenta',
      'vbCyan',
      'vbWhite',

      'vbUseCompareOption',
      'vbBinaryCompare',
      'vbTextCompare',
      'vbDatabaseCompare',

      'vbMonday',
      'vbTuesday',
      'vbWednesday',
      'vbThursday',
      'vbFriday',
      'vbSaturday',
      'vbSunday',
      'vbUseSystemDayOfWeek',

       // First Week of Year constants
      'vbUseSystem',
      'vbFirstJan1',
      'vbFirstFourDays',
      'vbFirstFullWeek',

      'vbGeneralDate',
      'vbLongDate',
      'vbShortDate',
      'vbLongTime',
      'vbShortTime',


      'vbKeyLButton',
      'vbKeyRButton',
      'vbKeyCancel',
      'vbKeyMButton',
      'vbKeyBack',
      'vbKeyTab',
      'vbKeyClear',
      'vbKeyReturn',
      'vbKeyShift',
      'vbKeyControl',
      'vbKeyMenu',
      'vbKeyPause',
      'vbKeyCapital',
      'vbKeyEscape',
      'vbKeySpace',
      'vbKeyPageUp',
      'vbKeyPageDown',
      'vbKeyEnd',
      'vbKeyHome',
      'vbKeyLeft',
      'vbKeyUp',
      'vbKeyRight',
      'vbKeyDown',
      'vbKeySelect',
      'vbKeyPrint',
      'vbKeyExecute',
      'vbKeySnapshot',
      'vbKeyInsert',
      'vbKeyDelete',
      'vbKeyHelp',
      'vbKeyNumlock',
      'vbKeyA',
      'vbKeyB',
      'vbKeyC',
      'vbKeyD',
      'vbKeyE',
      'vbKeyF',
      'vbKeyG',
      'vbKeyH',
      'vbKeyI',
      'vbKeyJ',
      'vbKeyK',
      'vbKeyL',
      'vbKeyM',
      'vbKeyN',
      'vbKeyO',
      'vbKeyP',
      'vbKeyQ',
      'vbKeyR',
      'vbKeyS',
      'vbKeyT',
      'vbKeyU',
      'vbKeyV',
      'vbKeyW',
      'vbKeyX',
      'vbKeyY',
      'vbKeyZ',
      'vbKey0',
      'vbKey1',
      'vbKey2',
      'vbKey3',
      'vbKey4',
      'vbKey5',
      'vbKey6',
      'vbKey7',
      'vbKey8',
      'vbKey9',
      'vbKeyNumpad0',
      'vbKeyNumpad1',
      'vbKeyNumpad2',
      'vbKeyNumpad3',
      'vbKeyNumpad4',
      'vbKeyNumpad5',
      'vbKeyNumpad6',
      'vbKeyNumpad7',
      'vbKeyNumpad8',
      'vbKeyNumpad9',
      'vbKeyMultiply',
      'vbKeyAdd',
      'vbKeySeparator',
      'vbKeySubtract',
      'vbKeyDecimal',
      'vbKeyDivide',
      'vbKeyF1',
      'vbKeyF2',
      'vbKeyF3',
      'vbKeyF4',
      'vbKeyF5',
      'vbKeyF6',
      'vbKeyF7',
      'vbKeyF8',
      'vbKeyF9',
      'vbKeyF10',
      'vbKeyF11',
      'vbKeyF12',
      'vbKeyF13',
      'vbKeyF14',
      'vbKeyF15',
      'vbKeyF16',

      'vbCrLf',
      'vbCr',
      'vbLf',
      'vbNewLine',
      'vbNullChar',
      'vbNullString',
      'vbObjectError',
      'vbTab',
      'vbBack',
      'vbFormFeed',
      'vbVerticalTab',
      

      // MsgBox constants
      'vbOKOnly',
      'vbOKCancel',
      'vbAbortRetryIgnore',
      'vbYesNoCancel',
      'vbYesNo',
      'vbRetryCancel',

      // MsgBox icon constants
      'vbCritical',
      'vbQuestion',
      'vbExclamation',
      'vbInformation',

      // MsgBox return values
      'vbOK',
      'vbCancel',
      'vbAbort',
      'vbRetry',
      'vbIgnore',
      'vbYes',
      'vbNo',

      'vbUpperCase',
      'vbLowerCase',
      'vbProperCase',
      'vbWide',
      'vbNarrow',
      'vbKatakana',
      'vbHiragana',
      'vbUnicode',
      'vbFromUnicode',

      'vbTrue',
      'vbFalse',
      'vbUseDefault',

      'vbEmpty',
      'vbNull',
      'vbInteger',
      'vbLong',
      'vbSingle',
      'vbDouble',
      'vbCurrency',
      'vbDate',
      'vbString',
      'vbObject',
      'vbError',
      'vbBoolean',
      'vbVariant',
      'vbDataObject',
      'vbDecimal',
      'vbByte',
      'vbUserDefinedType',
      'vbArray'


    ),

    // Object creation: New ClassName
    object_creation: $ => seq(
      token(/New/i),
      $.identifier
    ),

    // Parenthesized expressions: (expr)
    parenthesized_expression: $ => seq(
      '(',
      $.expression,
      ')'
    ),

    // Boolean literals: True/False
    boolean_literal: $ => choice(
      token(/True/i),
      token(/False/i)
    ),

    // Nothing literal
    nothing_literal: $ => token(/Nothing/i),

    // Bare function calls: Name(args)
    function_call: $ => prec.left(PREC.call, seq(
      $.identifier,
      $.argument_list
    )),

    // Enhanced property access: expr.Identifier (can be chained)
    property_access: $ => prec.left(PREC.member, seq(
      field('object', choice(
        $.identifier,
        $.property_access,
        $.indexed_access,
        $.function_call,
        $.parenthesized_expression
      )),
      '.',
      field('property', $.identifier)
    )),

    // Binary expressions with precedence
    binary_expression: $ => choice(
      prec.left(PREC.mul, seq($.expression, token(choice('*', '/')), $.expression)),
      prec.left(PREC.add, seq($.expression, token(choice('+', '-')), $.expression)),
      prec.left(PREC.concat, seq($.expression, token('&'), $.expression)),
      prec.left(PREC.relational, seq($.expression, token(choice('>=', '<=', '>', '<')), $.expression)),
      prec.left(PREC.equality, seq($.expression, token(choice('=', '<>')), $.expression))
    ),
    unary_expression: $ => seq(
      field('operator', choice('-', '+',$.keyword_Not)),
      field('argument', $.expression)
    ),

    // Terminals
    // Primitive types as per VBA
    primitive_type: $ => choice(
      'Boolean',
      'Byte',
      'Currency',
      'Date',
      'Decimal',
      'Double',
      'Integer',
      'Long',
      'LongLong',   // <-- added
      'Object' ,
      'Single',
      'String',
      'Variant'
    ), 

    // Your existing keyword tokens (keeping them all)
    keyword_Dim:     $ => token(/Dim/i),
    keyword_Const:   $ => token(/Const/i),
    keyword_As:      $ => token(/As/i),
    keyword_Global:  $ => token(/Global/i),
    keyword_Static:  $ => token(/Static/i),
    keyword_Sub:     $ => token(/Sub/i),
    keyword_Function:$ => token(/Function/i),
    keyword_End:     $ => token(/End/i),
    keyword_If:      $ => token(/If/i),
    keyword_Then:    $ => token(/Then/i),
    keyword_ElseIf: $ => choice(
      token(/ElseIf/i),
      // allow "Else If" on the same line (no newline allowed between)
      seq($.keyword_Else, token.immediate(/If/i))
    ),
    keyword_End_If: $ => seq(
      /[Ee][Nn][Dd]/,
      /[ \t]+/,
      /[Ii][Ff]/
    ),
    keyword_Else:    $ => token(/Else/i),
    keyword_Do:      $ => token(/Do/i),
    keyword_Loop:    $ => token(/Loop/i),
    keyword_While:   $ => token(/While/i),
    keyword_For:     $ => token(/For/i),
    keyword_To:      $ => token(/To/i),
    keyword_Next:    $ => token(/Next/i),
    keyword_Exit:    $ => token(/Exit/i),
    keyword_Select:  $ => token(/Select/i),
    keyword_Case:    $ => token(/Case/i),
    keyword_True:    $ => token(/True/i),
    keyword_False:   $ => token(/False/i),
    keyword_Set:     $ => token(/Set/i),
    keyword_Let:     $ => token(/Let/i),
    keyword_Call:    $ => token(/Call/i),
    keyword_With:    $ => token(/With/i),
    keyword_GoTo:    $ => token(/GoTo/i),
    keyword_On:      $ => token(/On/i),
    keyword_Error:   $ => token(/Error/i),
    keyword_Resume:  $ => token(/Resume/i),
    keyword_Nothing: $ => token(/Nothing/i),
    keyword_Not: $ => /[Nn][Oo][Tt]/,
    keyword_And:     $ => token(/And/i),
    keyword_Or:      $ => token(/Or/i),
    keyword_Mod:     $ => token(/Mod/i),
    keyword_Is:      $ => token(/Is/i),
    keyword_New:     $ => token(/New/i),
    keyword_Me:      $ => token(/Me/i),
    keyword_Option:  $ => token(/Option/i),
    keyword_Explicit:$ => token(/Explicit/i),
    keyword_ReDim:   $ => token(/ReDim/i),
    keyword_Preserve:$ => token(/Preserve/i),

    // Identifiers must not conflict with reserved keywords
    identifier: $ => token(prec(-1,
      seq(
        /[A-Za-z]/,
        repeat(/[A-Za-z0-9_]/)
      )
    )),

    integer_literal: _ => /\d+/,   
    byte_literal: $ => token(/\d{1,3}/),  // matches 0–255 in source       
    string_literal: $ => seq(
      '"',
      repeat(choice(
        token.immediate(prec(1, /[^"\r\n]+/)),  // Any chars except quote and newline
        '""'  // Escaped quote: "" in VBA means a literal "
      )),
      '"'
    ),
    number_literal: $ => token(choice(
      /\d*\.\d+([eE][+-]?\d+)?/,  // decimal or scientific float (e.g. 12.5, .3, 2.5e3)
      /\d+[eE][+-]?\d+/,          // scientific notation (e.g. 3E5)
      /\d+/                       // integer fallback
    )),
    currency_literal: $ => token(seq(
      optional("$"),
      /[0-9]+(\.[0-9]{1,4})?/
    )),

    date_literal: $ => seq(
      token.immediate('#'),
      // Simple (but practical) matcher: mm/dd/yyyy [hh:mm[:ss]]
      token.immediate(/[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{2,4}( [0-9]{1,2}:[0-9]{2}(:[0-9]{2})?)?/),
      token.immediate('#')
    ),
    float_literal: $ => token(/[0-9]+\.[0-9]+/),
    // VBA comment support
    comment: $ => token(seq(
      "'",
      /[^\r\n]*/
    )),

  }
});

/*
Enhanced Assignment Statement Features:

1. Simple variable assignment:
   x = 5
   name = "John"

2. Object property assignment:
   obj.Property = "value"
   Range("A1").Value = 100

3. Array/Collection assignment:
   arr(0) = "first"
   matrix(i, j) = calculation()
   dict("key") = value

4. Chained property assignment:
   worksheet.Cells(1, 1).Value = data
   obj.GetRange().Offset(1, 0).Value = result

5. Module-qualified assignment:
   Module1.GlobalVar = 10

6. Set statement for objects:
   Set obj = New Collection
   Set range = ActiveSheet.Range("A1")

7. Complex expressions on right side:
   x = obj.Method().Property + 5
   arr(i) = New CustomClass

The lvalue rule handles all valid left-hand side patterns while maintaining
compatibility with your existing expression system.
*/