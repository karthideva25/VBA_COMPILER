#include <tree_sitter/parser.h>
#include <wctype.h>
#include <stdbool.h>

enum TokenType {
  LINE_CONTINUATION,
};

void *tree_sitter_vba_external_scanner_create() { return NULL; }
void tree_sitter_vba_external_scanner_destroy(void *_p) {}
void tree_sitter_vba_external_scanner_reset(void *_p) {}
unsigned tree_sitter_vba_external_scanner_serialize(void *_p, char *_buffer) { return 0; }
void tree_sitter_vba_external_scanner_deserialize(void *_p, const char *_b, unsigned _n) {}

bool tree_sitter_vba_external_scanner_scan(void *_payload, TSLexer *lexer, const bool *valid_symbols) {
  if (valid_symbols[LINE_CONTINUATION]) {
    while (iswspace(lexer->lookahead)) {
      lexer->advance(lexer, true);
    }

    if (lexer->lookahead == '_') {
      lexer->advance(lexer, false);
      lexer->result_symbol = LINE_CONTINUATION;
      return true;
    }
  }

  return false;
}
