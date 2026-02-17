pub mod ast;
pub mod context;
pub mod interpreter;
pub mod runtime_config;
pub mod vm;
pub mod host;

pub use ast::{Program, Statement as VbaAstNode, build_ast as _build_ast};
pub use context::{Context, Value as VbaValue};
pub use runtime_config::{RuntimeConfig, RuntimeConfigBuilder};
pub use interpreter::execute_ast;
pub use vm::{ProgramExecutor, VbaRuntime};

use tree_sitter::TreeCursor;

/// Turn a `TreeCursor` at the root into a flat `Vec<Statement>` for your `main.rs`.
pub fn walk_parse_tree(cursor: &mut TreeCursor, source: &str) -> Vec<VbaAstNode> {
    let root = cursor.node();
    ast::build_ast(root, source).statements
}

/// Existing parse‐tree printer you already have…
use tree_sitter::Parser;
use vba_parser::language as vba_language;

pub fn print_parse_tree(code: &str) {
    let mut parser = Parser::new();
    parser.set_language(vba_language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    let root = tree.root_node();
    let mut cursor = root.walk();
    println!("🔍 Tree-sitter Parse Tree:");
    print_node(&mut cursor, code, 0);
}

fn print_node(cursor: &mut TreeCursor, src: &str, indent: usize) {
    loop {
        let n = cursor.node();
        let text = n.utf8_text(src.as_bytes()).unwrap_or("");
        println!("{:indent$}{}: {:?}", "", n.kind(), text, indent = indent*2);
        if cursor.goto_first_child() {
            print_node(cursor, src, indent+1);
            cursor.goto_parent();
        }
        if !cursor.goto_next_sibling() {
            break;
        }
    }
}
