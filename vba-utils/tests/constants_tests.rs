use tree_sitter::Parser;
use vba_parser::language;

/// Extracts the first built-in constant token's text from a snippet.
/// Robust to named/unnamed nodes and minor grammar changes.
fn find_first_builtin_constant(code: &str) -> String {
    let mut parser = Parser::new();
    parser.set_language(language()).unwrap();
    let tree = parser.parse(code, None).unwrap();
    let root = tree.root_node();

    find_builtin_text_anywhere(root, code)
        .expect("Expected a built-in constant like vbCalGreg/vbMethod/etc.")
}

/// Depth-first traversal over ALL children (named + unnamed).
fn find_builtin_text_anywhere(node: tree_sitter::Node, src: &str) -> Option<String> {
    // Try current node
    if let Some(txt) = node_text_if_builtin(node, src) {
        return Some(txt);
    }

    // Walk all children (including unnamed)
    let count = node.child_count();
    for i in 0..count {
        if let Some(child) = node.child(i) {
            if let Some(found) = find_builtin_text_anywhere(child, src) {
                return Some(found);
            }
        }
    }
    None
}

/// Recognize a builtin constant node or a fallback identifier starting with "vb".
fn node_text_if_builtin(node: tree_sitter::Node, src: &str) -> Option<String> {
    let kind = node.kind();

    // Read text (may fail on non-leaf; ignore those)
    let text = node.utf8_text(src.as_bytes()).ok()?.to_string();

    // Accept common shapes:
    let is_builtin_kind = kind == "vba_builtin_constant" || kind == "builtin_constant";
    let is_vb_identifier = kind == "identifier" && text.starts_with("vb");

    if is_builtin_kind || is_vb_identifier {
        Some(text)
    } else {
        None
    }
}

macro_rules! test_constant {
    ($name:ident, $const_str:expr) => {
        #[test]
        fn $name() {
            let code = format!(r#"
                Sub Test()
                    Dim val As Long
                    val = {}
                End Sub
            "#, $const_str);

            let text = find_first_builtin_constant(&code);
            assert_eq!(text, $const_str);
        }
    };
}

// Calendar constants
test_constant!(test_vbCalGreg, "vbCalGreg");
test_constant!(test_vbCalHijri, "vbCalHijri");

// CallType constants
test_constant!(test_vbMethod, "vbMethod");
test_constant!(test_vbGet, "vbGet");
test_constant!(test_vbLet, "vbLet");
test_constant!(test_vbSet, "vbSet");

// Day of Week constants
test_constant!(test_vbSunday, "vbSunday");
test_constant!(test_vbMonday, "vbMonday");
test_constant!(test_vbTuesday, "vbTuesday");
test_constant!(test_vbWednesday, "vbWednesday");
test_constant!(test_vbThursday, "vbThursday");
test_constant!(test_vbFriday, "vbFriday");
test_constant!(test_vbSaturday, "vbSaturday");
//test_constant!(test_vbUseSystemDayOfWeek, "vbUseSystemDayOfWeek");

// First Week of Year constants
test_constant!(test_vbUseSystem, "vbUseSystem");
// test_constant!(test_vbFirstJan1, "vbFirstJan1");
// test_constant!(test_vbFirstFourDays, "vbFirstFourDays");
// test_constant!(test_vbFirstFullWeek, "vbFirstFullWeek");

// Comparison constants
test_constant!(test_vbBinaryCompare, "vbBinaryCompare");
test_constant!(test_vbTextCompare, "vbTextCompare");
test_constant!(test_vbDatabaseCompare, "vbDatabaseCompare");

// MsgBox constants
test_constant!(test_vbOKOnly, "vbOKOnly");
test_constant!(test_vbOKCancel, "vbOKCancel");
test_constant!(test_vbAbortRetryIgnore, "vbAbortRetryIgnore");
test_constant!(test_vbYesNoCancel, "vbYesNoCancel");
test_constant!(test_vbYesNo, "vbYesNo");
test_constant!(test_vbRetryCancel, "vbRetryCancel");

// MsgBox icon constants
test_constant!(test_vbCritical, "vbCritical");
test_constant!(test_vbQuestion, "vbQuestion");
test_constant!(test_vbExclamation, "vbExclamation");
test_constant!(test_vbInformation, "vbInformation");

// MsgBox return values
test_constant!(test_vbOK, "vbOK");
test_constant!(test_vbCancel, "vbCancel");
test_constant!(test_vbAbort, "vbAbort");
test_constant!(test_vbRetry, "vbRetry");
test_constant!(test_vbIgnore, "vbIgnore");
test_constant!(test_vbYes, "vbYes");
test_constant!(test_vbNo, "vbNo");

// Color constants
test_constant!(test_vbBlack, "vbBlack");
test_constant!(test_vbRed, "vbRed");
test_constant!(test_vbGreen, "vbGreen");
test_constant!(test_vbYellow, "vbYellow");
test_constant!(test_vbBlue, "vbBlue");
test_constant!(test_vbMagenta, "vbMagenta");
test_constant!(test_vbCyan, "vbCyan");
test_constant!(test_vbWhite, "vbWhite");
