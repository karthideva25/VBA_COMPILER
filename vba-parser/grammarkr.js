// Tree‑sitter grammar for VBA / VB6 – **pass 3**
// ==================================================
// Adds the remaining *high‑impact* production rules:
//   • File‑I/O statements → `Open`, `Close`, `Print`, `Input`, `Line Input`,
//     `Write`, `Get`, `Put`, `Seek`, `Kill`, `ChDir`, `ChDrive`, `MkDir`, `RmDir`
//   • Pre‑compiler / macro directives → `#CONST`, `#IF … THEN`, `#ELSE`,
//     `#ELSEIF`, `#END IF`
//   • Generalised `file_number` helper.
//   • Expanded `block_stmt` + `module_declaration_el` to recognise the above.
//
// The file‑I/O rules are *structurally correct* for most real‑world VBA but
// do not yet capture every optional clause (e.g. full `ACCESS`/`LOCK` matrix).
// They’re easy to extend—look for `TODO` comments.
//
// Next pass (if you need it) can refine error‑handling (`On Error`, `Resume`),
// `Declare` PtrSafe alias nuances, and full pre‑compiler nesting.
//
// ───────────────────────────────────────────────────────────────────────────

const PREC = {
  assign:       15,
  pow:          14,
  unary:        13,
  multiplicative:12,
  mod:          11,
  additive:     10,
  concat:        9,
  relational:    8,
  not:           7,
  and:           6,
  or_xor:        5,
  eqv:           4,
  imp:           3,
};

function commaSep1(rule) { return seq(rule, repeat(seq(',', rule))); }
function sep1(rule, sep) { return seq(rule, repeat(seq(sep, rule))); }

module.exports = grammar({
  name: 'vba',

  externals: $ => [ $.line_continuation ],

  extras: $ => [ /[ \t\f\v]+/, /[\r\n]+/, $.comment ],

  word: $ => $.identifier,

  rules: {
    // ───────────────────────────── TOP LEVEL ──
    // source_file: $ => seq(optional($.module), repeat($.end_of_line)),
    source_file: $ => seq(optional($.module), repeat($.end_of_line)),


    module: $ => seq(
      optional($.module_header), repeat($.end_of_line),
      optional($.module_config), repeat($.end_of_line),
      optional($.module_attributes), repeat($.end_of_line),
      optional($.module_declarations), repeat($.end_of_line),
      $.module_body,   // not optional anymore
    ),
    guid: _ => /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/,

    module_header: $ => seq('VERSION', $.double_literal, optional('CLASS')),
    module_config: $ => seq('BEGIN', optional(seq($.guid, $.identifier)), repeat($.end_of_line), repeat1($.module_config_element), 'END'),
    module_config_element: $ => seq(field('key', $.identifier), '=', field('value', $.literal), optional(seq(':', $.literal)), repeat($.end_of_line)),
    module_attributes: $ => repeat1(seq($.attribute_stmt, repeat1($.end_of_line))),

   // module_declarations: $ => seq($.module_declaration_el, repeat(seq(repeat1($.end_of_line), $.module_declaration_el)), repeat($.end_of_line)),
    module_declarations: $ => repeat1(seq($.module_declaration_el, repeat1($.end_of_line))),
    module_declaration_el: $ => choice(
      $.comment,
      $.macro_stmt,              // <── added
      $.const_stmt,
      $.declare_stmt,
      $.variable_stmt,
      $.module_option,
    ),
    module_option: $ => choice(
      seq('OPTION', 'BASE', $.integer_literal),
      seq('OPTION', 'COMPARE', choice('BINARY', 'TEXT', 'DATABASE')),
      'OPTION EXPLICIT',
      seq('OPTION', 'PRIVATE', 'MODULE')
    ),

    module_body: $ => seq($.module_body_el, repeat(seq(repeat1($.end_of_line), $.module_body_el)), repeat($.end_of_line)),
    module_body_el: $ => choice(
      $.sub_stmt,
      $.function_stmt,
      $.property_get_stmt,
      $.property_set_stmt,
      $.property_let_stmt,
      $.macro_stmt,              // allow #IF blocks in body as well
    ),

    // ───────────────────── STATEMENTS (declarations) ──
    const_stmt: $ => seq(optional($.visibility), 'CONST', commaSep1($.const_binding)),
    const_binding: $ => seq(field('name', $.identifier), optional($.type_hint), optional($.as_type_clause), '=', field('value', $.value_stmt)),

    declare_stmt: $ => seq(optional($.visibility), 'DECLARE', optional('PTRSAFE'), choice(seq('FUNCTION', optional($.type_hint)), 'SUB'), field('name', $.identifier), optional($.type_hint), 'LIB', $.string_literal, optional(seq('ALIAS', $.string_literal)), optional($.arg_list), optional($.as_type_clause)),

    variable_stmt: $ => seq(choice('DIM', 'STATIC', $.visibility), repeat1($.variable_binding)),
    variable_binding: $ => seq($.identifier, optional(seq('(', optional($.subscripts), ')')), optional($.type_hint), optional($.as_type_clause)),

    // ───────────────────── PROCEDURE HEADS ── (unchanged)
    sub_stmt: $ => seq(optional($.visibility), optional('STATIC'), 'SUB', field('name', $.identifier), optional($.arg_list), $.end_of_statement, optional($.block), 'END', 'SUB'),
    function_stmt: $ => seq(optional($.visibility), optional('STATIC'), 'FUNCTION', field('name', $.identifier), optional($.type_hint), optional($.arg_list), optional($.as_type_clause), $.end_of_statement, optional($.block), 'END', 'FUNCTION'),
    property_get_stmt: $ => seq(optional($.visibility), optional('STATIC'), 'PROPERTY', 'GET', field('name', $.identifier), optional($.type_hint), optional($.arg_list), optional($.as_type_clause), $.end_of_statement, optional($.block), 'END', 'PROPERTY'),
    property_set_stmt: $ => seq(optional($.visibility), optional('STATIC'), 'PROPERTY', 'SET', field('name', $.identifier), optional($.arg_list), $.end_of_statement, optional($.block), 'END', 'PROPERTY'),
    property_let_stmt: $ => seq(optional($.visibility), optional('STATIC'), 'PROPERTY', 'LET', field('name', $.identifier), optional($.arg_list), $.end_of_statement, optional($.block), 'END', 'PROPERTY'),

    // ────────────────────── BLOCK & FLOW ──
    attribute_stmt: $ => seq('ATTRIBUTE', $.identifier, '=', $.literal, repeat(seq(',', $.literal)) ),

    block: $ => seq(repeat1(seq($.block_stmt, $.end_of_statement))),
    block_stmt: $ => choice(
      // existing
      $.const_stmt,
      $.variable_stmt,
      $.let_stmt,
      $.exit_stmt,
      $.return_stmt,
      $.stop_stmt,
      $.if_stmt,
      $.select_case_stmt,
      $.for_next_stmt,
      $.for_each_stmt,
      $.do_loop_stmt,
      $.while_wend_stmt,
      $.with_stmt,
      // new
      $.file_io_stmt,
      $.macro_stmt,
      $.implicit_call_stmt,
    ),

    let_stmt:   $ => seq(optional('LET'), $.implicit_call_stmt, '=', $.value_stmt),
    exit_stmt:  $ => seq('EXIT', choice('SUB', 'FUNCTION', 'PROPERTY', 'DO', 'FOR')),
    return_stmt:$ => 'RETURN',
    stop_stmt:  $ => 'STOP',

    // ────── IF / ELSE (unchanged) ──────
    if_stmt: $ => choice(
      seq('IF', $.value_stmt, 'THEN', $.block_stmt, optional(seq('ELSE', $.block_stmt))),
      seq($.if_block, repeat($.elseif_block), optional($.else_block), 'END', 'IF')
    ),
    if_block: $ => seq('IF', $.value_stmt, 'THEN', $.end_of_statement, optional($.block)),
    elseif_block: $ => seq('ELSEIF', $.value_stmt, 'THEN', $.end_of_statement, optional($.block)),
    else_block: $ => seq('ELSE', $.end_of_statement, optional($.block)),

    // ────── SELECT CASE / LOOPS / WITH (unchanged) ──────
    select_case_stmt: $ => seq('SELECT', 'CASE', $.value_stmt, $.end_of_statement, repeat($.case_clause), 'END', 'SELECT'),
    case_clause: $ => seq('CASE', choice('ELSE', sep1($.case_cond, ',')), $.end_of_statement, optional($.block)),
    case_cond: $ => choice(seq('IS', $.comparison_op, $.value_stmt), seq($.value_stmt, 'TO', $.value_stmt), $.value_stmt),

    for_next_stmt: $ => seq('FOR', $.identifier, '=', $.value_stmt, 'TO', $.value_stmt, optional(seq('STEP', $.value_stmt)), $.end_of_statement, optional($.block), 'NEXT', optional($.identifier)),
    for_each_stmt: $ => seq('FOR', 'EACH', $.identifier, 'IN', $.value_stmt, $.end_of_statement, optional($.block), 'NEXT', optional($.identifier)),
    do_loop_stmt: $ => choice(
      seq('DO', $.end_of_statement, optional($.block), 'LOOP'),
      seq('DO', choice('WHILE', 'UNTIL'), $.value_stmt, $.end_of_statement, optional($.block), 'LOOP'),
      seq('DO', $.end_of_statement, optional($.block), 'LOOP', choice('WHILE', 'UNTIL'), $.value_stmt)
    ),
    while_wend_stmt: $ => seq('WHILE', $.value_stmt, $.end_of_statement, optional($.block), 'WEND'),
    with_stmt: $ => seq('WITH', choice($.implicit_call_stmt, seq('NEW', $.type_)), $.end_of_statement, optional($.block), 'END', 'WITH'),

    // ────── FILE‑I/O STATEMENTS ──────
    file_io_stmt: $ => choice(
      $.open_stmt,
      $.close_stmt,
      $.print_stmt,
      $.input_stmt,
      $.line_input_stmt,
      $.write_stmt,
      $.get_stmt,
      $.put_stmt,
      $.seek_stmt,
      $.kill_stmt,
      $.mkdir_stmt,
      $.rmdir_stmt,
      $.chdir_stmt,
      $.chdrive_stmt,
    ),

    close_stmt: $ => seq('CLOSE', optional(seq($.file_number, repeat(seq(',', $.file_number))))),

    open_stmt: $ => seq('OPEN', $.value_stmt, 'FOR', choice('APPEND', 'BINARY', 'INPUT', 'OUTPUT', 'RANDOM'), optional(seq('ACCESS', choice('READ', 'WRITE', seq('READ', 'WRITE')))), optional(seq(choice('SHARED', seq('LOCK', choice('READ', 'WRITE', seq('READ', 'WRITE')))))), 'AS', $.file_number, optional(seq('LEN', '=', $.value_stmt)) ),

    print_stmt: $ => seq('PRINT', $.file_number, ',', optional($.output_list)),
    output_list: $ => sep1($.value_stmt, choice(',', ';')),

    input_stmt: $ => seq('INPUT', $.file_number, repeat1(seq(',', $.value_stmt))),
    line_input_stmt: $ => seq('LINE', 'INPUT', $.file_number, ',', $.value_stmt),

    write_stmt: $ => seq('WRITE', $.file_number, ',', optional($.output_list)),

    get_stmt: $ => seq('GET', $.file_number, ',', optional($.value_stmt), ',', $.value_stmt),
    put_stmt: $ => seq('PUT', $.file_number, ',', optional($.value_stmt), ',', $.value_stmt),

    seek_stmt: $ => seq('SEEK', $.file_number, ',', $.value_stmt),

    kill_stmt:  $ => seq('KILL', $.value_stmt),
    mkdir_stmt: $ => seq('MKDIR', $.value_stmt),
    rmdir_stmt: $ => seq('RMDIR', $.value_stmt),
    chdir_stmt: $ => seq('CHDIR', $.value_stmt),
    chdrive_stmt: $ => seq('CHDRIVE', $.value_stmt),

    file_number: $ => seq(optional('#'), $.value_stmt),

    // ────── MACRO / PRE‑COMPILER ──────
    macro_stmt: $ => choice($.macro_const_stmt, $.macro_if_block),

    macro_const_stmt: $ => seq('#CONST', $.identifier, '=', $.value_stmt),

    macro_if_block: $ => seq(
      '#IF', $.value_stmt, 'THEN', repeat($.macro_body_line),
      repeat($.macro_elseif_block),
      optional($.macro_else_block),
      '#END', 'IF'
    ),
    macro_elseif_block: $ => seq('#ELSEIF', $.value_stmt, 'THEN', repeat($.macro_body_line)),
    macro_else_block: $ => seq('#ELSE', repeat($.macro_body_line)),

    macro_body_line: $ => token(/[^\r\n]+/),

    // ─────────────────────── EXPRESSIONS (unchanged) ──
    value_stmt: $ => choice(
      prec.right(PREC.assign, seq($.implicit_call_stmt, ':=', $.value_stmt)),
      prec.right(PREC.pow, seq($.value_stmt, '^', $.value_stmt)),
      prec(PREC.unary, seq('-', $.value_stmt)),
      prec(PREC.unary, seq('+', $.value_stmt)),
      prec.left(PREC.multiplicative, seq($.value_stmt, choice('*', '/', '\\'), $.value_stmt)),
      prec.left(PREC.mod, seq($.value_stmt, 'MOD', $.value_stmt)),
      prec.left(PREC.additive, seq($.value_stmt, choice('+', '-'), $.value_stmt)),
      prec.left(PREC.concat, seq($.value_stmt, '&', $.value_stmt)),
      prec.left(PREC.relational, seq($.value_stmt, $.comparison_op, $.value_stmt)),
      prec.right(PREC.not, seq('NOT', $.value_stmt)),
      prec.left(PREC.and, seq($.value_stmt, 'AND', $.value_stmt)),
      prec.left(PREC.or_xor, seq($.value_stmt, choice('OR', 'XOR'), $.value_stmt)),
      prec.left(PREC.eqv, seq($.value_stmt, 'EQV', $.value_stmt)),
      prec.left(PREC.imp, seq($.value_stmt, 'IMP', $.value_stmt)),
      $.literal,
      $.implicit_call_stmt,
      seq('(', sep1($.value_stmt, ','), ')'),
      seq('NEW', $.value_stmt),
    ),

    comparison_op: _ => choice('<', '<=', '>', '>=', '=', '<>', 'IS', 'LIKE'),

    // ───────────────────── CALLS & ATOMS (unchanged) ──
    implicit_call_stmt: $ => seq($.identifier, optional($.type_hint), optional(seq('(', optional(sep1($.args_call, choice(',', ';'))), ')')), optional($.dictionary_call), repeat(seq('(', $.subscripts, ')'))),
    args_call: $ => $.value_stmt,
    dictionary_call: $ => seq('!', $.identifier),
    arg_list: $ => seq('(', optional(sep1($.arg_decl, ',')), ')'),
    arg_decl: $ => seq(optional('OPTIONAL'), optional(choice('BYVAL', 'BYREF')), optional('PARAMARRAY'), $.identifier, optional($.type_hint), optional($.as_type_clause), optional(seq('=', $.value_stmt))),
    as_type_clause: $ => seq('AS', optional('NEW'), $.type_),
    type_: $ => choice($.base_type, $.complex_type, seq($.base_type, '(', ')')),
    base_type: $ => choice('BOOLEAN', 'BYTE', 'DATE', 'DOUBLE', 'INTEGER', 'LONG', 'SINGLE', seq('STRING', optional(seq('*', $.integer_literal))), 'VARIANT', 'COLLECTION'),
    complex_type: $ => seq($.identifier, repeat(seq(choice('.', '!'), $.identifier))),
    subscripts: $ => sep1($.	value_stmt, ','),

    // ─────────────────────── LEXICAL (unchanged) ──
    identifier: _ => /[A-Za-z_\p{L}][A-Za-z0-9_\p{L}]*/i,
    type_hint: _ => token(/[%&$#!@]/),
    integer_literal: _ => /[0-9]+/,                // TODO: octal & hex
    double_literal:  _ => /[0-9]*\.[0-9]+([eE][+-]?[0-9]+)?/,
    string_literal:  _ => /"([^"\n]|""\n?)*"/,
    date_literal:    _ => /#[^#\n]+#/,
    literal: $ => choice($.integer_literal, $.double_literal, $.string_literal, $.date_literal, 'TRUE', 'FALSE', 'NULL', 'NOTHING'),

    visibility: _ => choice('PUBLIC', 'PRIVATE', 'FRIEND', 'GLOBAL'),
    comment: _ => token(/'.*/),
    end_of_line: _ => /[\r\n]+/,
    line_continuation: _ => /_\r?\n/,
    end_of_statement: _ => choice(
        /\r?\n/,
        ':',
      ),
      

  },
});
