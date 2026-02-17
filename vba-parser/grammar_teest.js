module.exports = grammar({
    name: 'vba_test',
    rules: {
      source_file: $ => repeat($.statement),
      statement: $ => /.+/,
    }
  });
  