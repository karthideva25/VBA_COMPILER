const PREC = {
  call: 16,
  assign: 15,
  pow: 14,
  unary: 13,
  multiplicative: 12,
  additive: 11,
  concat: 10,
  relational: 9,
  equality: 8,
  and: 7,
  or: 6,
  xor: 5,
};

// Helper for comma-separated lists
function commaSep1(rule) {
  return seq(rule, repeat(seq(',', rule)));
}

module.exports = grammar({
  name: 'vba',

  extras: $ => [
    /[ \t\r]+/,         // spaces, tabs, carriage returns
    $.line_continuation,  // underscore line continuation
  ],

  externals: $ => [
    $.line_continuation,
  ],

  conflicts: $ => [
    [$.member_expression, $.call_expression],
    [$.call_statement, $.expression_statement],
    [$.assignment_statement, $.expression],
  ],

  precedences: $ => [
    [$.call_expression],    // highest precedence for calls
    [$.concat_expression],  // concatenation
    [$.assignment_expression],
  ],

  rules: {
    source_file: $ => seq(
      repeat1($.module),
      optional($.end_of_line)
    ),

    module: $ => prec.right(seq(
      optional($.module_header),
      optional($.module_option),
      repeat($.module_attribute),
      repeat1($.member)
    )),

    module_header: $ => seq('VERSION', $.string_literal, $.end_of_line),
    module_option: $ => seq('OPTION', choice('EXPLICIT', 'BASE'), $.end_of_line),
    module_attribute: $ => seq('ATTRIBUTE', $.identifier, '=', $.literal, $.end_of_line),

    member: $ => choice(
      $.subroutine,
      $.function_statement,
      $.const_statement,
      $.dim_statement
    ),

    subroutine: $ => seq(
      optional($.visibility),
      optional('STATIC'),
      'SUB', $.identifier,
      optional($.parameter_list),
      $.end_of_statement,
      repeat($.statement),
      'END', 'SUB', $.end_of_statement
    ),

    function_statement: $ => seq(
      optional($.visibility),
      optional('STATIC'),
      'FUNCTION', $.identifier,
      optional($.parameter_list),
      optional($.as_clause),
      $.end_of_statement,
      repeat($.statement),
      'END', 'FUNCTION', $.end_of_statement
    ),

    const_statement: $ => seq(
      optional($.visibility),
      'CONST', commaSep1($.const_binding),
      $.end_of_statement
    ),

    const_binding: $ => seq($.identifier, '=', $.expression),

    dim_statement: $ => seq(
      optional($.visibility),
      'DIM', commaSep1($.dim_binding),
      $.end_of_statement
    ),

    dim_binding: $ => seq(
      $.identifier,
      optional($.as_clause),
      optional(seq('=', $.expression))
    ),

    parameter_list: $ => seq('(', optional(commaSep1($.parameter)), ')'),
    parameter: $ => seq(
      optional(choice('BYVAL', 'BYREF')),
      $.identifier,
      optional($.as_clause),
      optional(seq('=', $.expression))
    ),

    as_clause: $ => seq('AS', $.type_name),

    statement: $ => choice(
      $.empty_statement,
      $.subroutine,
      $.function_statement,
      $.dim_statement,
      $.const_statement,
      $.assignment_statement,
      $.call_statement,
      $.expression_statement,
      $.if_statement,
      $.for_statement,
      $.while_statement,
      $.exit_statement
    ),

    empty_statement: $ => $.end_of_statement,

    assignment_statement: $ => seq(
      choice($.member_expression, $.identifier),
      '=',
      $.expression,
      $.end_of_statement
    ),

    call_statement: $ => seq(
      $.call_expression,
      $.end_of_statement
    ),

    expression_statement: $ => seq(
      choice(
        $.binary_expression,
        $.comparison_expression,
        $.concat_expression,
        $.unary_expression,
        $.primary_expression,
        $.member_expression
      ),
      $.end_of_statement
    ),

    if_statement: $ => seq(
      'IF', $.expression, 'THEN', $.end_of_statement,
      repeat($.statement),
      optional(seq('ELSE', repeat($.statement))),
      'END', 'IF', $.end_of_statement
    ),

    for_statement: $ => seq(
      'FOR', $.identifier, '=', $.expression, 'TO', $.expression,
      optional(seq('STEP', $.expression)), $.end_of_statement,
      repeat($.statement),
      'NEXT', optional($.identifier), $.end_of_statement
    ),

    while_statement: $ => seq(
      'WHILE', $.expression, $.end_of_statement,
      repeat($.statement),
      'WEND', $.end_of_statement
    ),

    exit_statement: $ => seq(
      'EXIT', choice('SUB','FUNCTION','FOR','WHILE'),
      $.end_of_statement
    ),

    expression: $ => choice(
      $.assignment_expression,
      $.binary_expression,
      $.comparison_expression,
      $.concat_expression,
      $.unary_expression,
      $.call_expression,
      $.member_expression,
      $.primary_expression
    ),

    assignment_expression: $ => prec.right(PREC.assign, seq(
      choice($.member_expression, $.identifier),
      '=',
      $.expression
    )),
    binary_expression: $ => prec.left(PREC.additive, seq(
      $.expression, choice('+','-','*','/','^'), $.expression
    )),
    comparison_expression: $ => prec.left(PREC.relational, seq(
      $.expression, choice('=','<>','<','<=','>','>='), $.expression
    )),
    concat_expression: $ => prec.left(PREC.concat, seq(
      $.expression, '&', $.expression
    )),
    unary_expression: $ => prec(PREC.unary, seq(
      choice('+','-','NOT'), $.expression
    )),

    primary_expression: $ => choice(
      $.number_literal,
      $.string_literal,
      $.boolean_literal,
      $.identifier,
      $.parenthesized_expression
    ),

    parenthesized_expression: $ => seq('(', $.expression, ')'),

    member_expression: $ => prec.left(PREC.concat, seq(
      choice($.call_expression, $.member_expression, $.identifier),
      '.',
      $.identifier
    )),
    call_expression: $ => prec(PREC.call, seq(
      choice($.member_expression, $.identifier),
      $.argument_list
    )),
    argument_list: $ => seq('(', optional(commaSep1($.expression)), ')'),

    number_literal: _ => /\d+(\.\d+)?([eE][+-]?\d+)?/,
    string_literal: _ => token(seq(
      '"', repeat(choice(/[^"\n]/, '""')), '"'
    )),
    boolean_literal: _ => choice('TRUE','FALSE'),

    // Generic literal for attributes and consts
    literal: $ => choice(
      $.number_literal,
      $.string_literal,
      $.boolean_literal
    ),

    identifier: _ => /[A-Za-z_][A-Za-z0-9_]*/,
    type_name: $ => $.identifier,
    visibility: _ => choice('PUBLIC','PRIVATE'),

    end_of_statement: _ => token(choice(':', /\r?\n/)),
    end_of_line: _ => /\r?\n/,
    line_continuation: _ => '_',
  }
});
