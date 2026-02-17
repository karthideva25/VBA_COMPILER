package tree_sitter_vbatestjune_test

import (
	"testing"

	tree_sitter "github.com/tree-sitter/go-tree-sitter"
	tree_sitter_vbatestjune "github.com/tree-sitter/tree-sitter-vbatestjune/bindings/go"
)

func TestCanLoadGrammar(t *testing.T) {
	language := tree_sitter.NewLanguage(tree_sitter_vbatestjune.Language())
	if language == nil {
		t.Errorf("Error loading Vbatestjune grammar")
	}
}
