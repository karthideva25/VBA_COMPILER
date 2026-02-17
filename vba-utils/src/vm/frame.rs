use crate::ast::{Statement, DoWhileStatement};

/// A single execution frame (analogous to a call stack frame in a real VM).
#[derive(Debug, Clone)]
pub struct Frame {
    pub id: usize,                      // Unique frame ID for debugging
    pub kind: FrameKind,                // What type of frame is this?
    pub list_id: usize,                 // Statement list ID (for resume tracking)
    pub pc: usize,                      // Program counter within the list
    pub statements: Vec<Statement>,     // The statements in this frame
    pub depth: usize,                   // Nesting depth
}

/// Different types of frames (each has different semantics for control flow).
#[derive(Debug, Clone)]
pub enum FrameKind {
    Main,                               // Top-level sub body
    For {
        counter: String,
        current_value: i64,
        end_value: i64,
        step: i64,
    },
    Do {
        statement: DoWhileStatement,    // Store complete statement to evaluate condition
        first_iteration: bool,          // Track if this is the first iteration
    },
    If,                                 // If/ElseIf/Else block
    Block,                              // Generic statement list (Call body, Type definition, etc.)
    With,                               // With block (object reference on context's with_stack)
}

impl Frame {
    pub fn new(
        id: usize,
        kind: FrameKind,
        list_id: usize,
        statements: Vec<Statement>,
        depth: usize,
    ) -> Self {
        Frame {
            id,
            kind,
            list_id,
            pc: 0,
            statements,
            depth,
        }
    }

    /// Advance to next statement.
    pub fn advance(&mut self) {
        self.pc += 1;
    }

    /// Jump to a specific PC.
    pub fn jump_to(&mut self, pc: usize) {
        self.pc = pc;
    }

    /// Is this frame finished (PC past end of statements)?
    pub fn is_done(&self) -> bool {
        self.pc >= self.statements.len()
    }

    /// Current statement (if any).
    pub fn current_statement(&self) -> Option<&Statement> {
        self.statements.get(self.pc)
    }
}
