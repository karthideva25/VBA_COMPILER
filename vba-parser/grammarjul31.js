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
//correct grammer but parsing everythis as error except striing literal
// Helper for comma-separated lists
function sep1(rule, sep) {
  return seq(rule, repeat(seq(sep, rule)));
}
function commaSep1(rule) {
  return sep1(rule, ',');
}

module.exports = grammar({
  name: 'vba',

  extras: $ => [
    /\s+/,                // whitespace
    $.line_continuation,   // line continuation
  ],

  externals: $ => [
    $.line_continuation,
  ],

  conflicts: $ => [
    [$.member_expression, $.call_expression],
    [$.module, $.module_declaration],
    [$.call_statement, $.expression_statement],
    [$.call_statement, $.expression],
  ],

  precedences: $ => [
    [$.call_expression],    // highest precedence: function/method calls
    [$.concat_expression],  // string concatenation
    [$.assignment_expression],
  ],

  rules: {
    source_file: $ => seq(
      repeat1($.module),
      optional($.end_of_line)
    ),

    module: $ => prec.right(seq(
      optional($.module_header),
      optional($.module_config),
      optional($.module_attributes),
      repeat1($.module_declaration)
    )),

    module_header: $ => seq('VERSION', $.string, $.end_of_line),

    module_config: $ => seq(
      'BEGIN', alias(/[A-Z_]+/, $.config_name), $.end_of_line,
      repeat($.statement),
      'END', alias(/[A-Z_]+/, $.config_name), $.end_of_line
    ),

    module_attributes: $ => repeat1(seq(
      'ATTRIBUTE', $.identifier, '=', $.literal, $.end_of_line
    )),

    module_declaration: $ => choice(
      $.option_statement,
      $.const_statement,
      $.variable_statement,
      $.subroutine,
      $.function_statement
    ),

    statement: $ => choice(
      $.empty_statement,
      $.assignment_statement,
      $.expression_statement,
      $.const_statement,
      $.variable_statement,
      $.option_statement,
      $.call_statement,
      $.subroutine,
      $.function_statement
    ),

    empty_statement: $ => $.end_of_statement,

    expression_statement: $ => seq(
      choice(
        $.binary_expression,
        $.comparison_expression,
        $.concat_expression,
        $.unary_expression,
        $.primary_expression
      ),
      $.end_of_statement
    ),

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

    subroutine: $ => seq(
      optional($.visibility),
      optional('STATIC'),
      'SUB',
      $.identifier,
      optional($.arg_list),
      $.end_of_statement,
      repeat($.statement),
      'END', 'SUB',
      $.end_of_statement
    ),

    function_statement: $ => seq(
      optional($.visibility),
      optional('STATIC'),
      'FUNCTION',
      $.identifier,
      optional($.arg_list),
      optional($.as_type),
      $.end_of_statement,
      repeat($.statement),
      'END', 'FUNCTION',
      $.end_of_statement
    ),

    option_statement: $ => seq(
      'OPTION', choice('EXPLICIT', 'BASE'),
      $.end_of_statement
    ),

    const_statement: $ => seq(
      optional($.visibility),
      'CONST',
      commaSep1($.const_binding),
      $.end_of_statement
    ),

    const_binding: $ => seq(
      $.identifier,
      '=',
      $.literal
    ),

    variable_statement: $ => seq(
      optional($.visibility),
      'DIM',
      commaSep1($.variable_declaration),
      $.end_of_statement
    ),

    variable_declaration: $ => seq(
      $.identifier,
      optional($.as_type),
      optional(seq('=', $.expression))
    ),

    as_type: $ => seq('AS', $.identifier),

    arg_list: $ => seq('(', optional(commaSep1($.arg)), ')'),

    arg: $ => seq(
      optional(choice('BYVAL', 'BYREF')),
      $.identifier,
      optional($.as_type),
      optional(seq('=', $.expression))
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
      $.expression,
      choice('+', '-', '*', '/', '^'),
      $.expression
    )),

    comparison_expression: $ => prec.left(PREC.relational, seq(
      $.expression,
      choice('=', '<>', '<', '<=', '>', '>='),
      $.expression
    )),

    concat_expression: $ => prec.left(PREC.concat, seq(
      $.expression,
      '&',
      $.expression
    )),

    unary_expression: $ => prec(PREC.unary, seq(
      choice('+', '-', 'NOT'),
      $.expression
    )),

    primary_expression: $ => choice(
      $.literal,
      $.identifier,
      $.parenthesized_expression
    ),

    parenthesized_expression: $ => seq('(', $.expression, ')'),

    member_expression: $ => prec.left(PREC.concat, seq(
      choice($.member_expression, $.call_expression, $.identifier),
      '.',
      $.identifier
    )),

    call_expression: $ => prec(PREC.call, seq(
      choice($.member_expression, $.identifier),
      $.arg_list
    )),

    literal: $ => choice(
      $.number,
      $.string,
      $.boolean
    ),
    number: _ => /\d+(\.\d+)?([eE][+-]?\d+)?/,
    string: _ => token(seq('"', repeat(choice(/[^"\n]/, '""')), '"')),
    boolean: _ => choice('TRUE', 'FALSE'),

    identifier: _ => /[A-Za-z_][A-Za-z0-9_]*/,  
    visibility: _ => choice('PUBLIC', 'PRIVATE'),

    end_of_statement: _ => token(choice(':', /\r?\n/)),
    end_of_line: _ => /\r?\n/,

    line_continuation: _ => '_',
  }
});
