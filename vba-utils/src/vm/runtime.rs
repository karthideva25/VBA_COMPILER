
use crate::ast::Statement;
use crate::ast::Expression;
use crate::context::Context;
use crate::interpreter::builtins::handle_builtin_call_bool;
use crate::context::ScopeKind;
use crate::interpreter::ControlFlow;
use std::collections::VecDeque;
use super::frame::{Frame, FrameKind};

/// The VBA execution virtual machine.
/// Maintains an explicit frame stack instead of relying on Rust's call stack.
pub struct VbaVm {
    frames: VecDeque<Frame>,           // Execution frame stack (bottom = main, top = current)
    next_frame_id: usize,
    pub vm_state: VmState,             // Current execution state
    pub saved_error_frame: Option<Frame>,
}

/// Execution state of the VM.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum VmState {
    Running,
    ErrorInProgress {
        error_label: String,
        from_frame_id: usize,          // Which frame did the error originate from?
    },
    ResumingFromError {
        target_frame_id: usize,
        target_pc: usize,
    },
    Halted,
}

impl VbaVm {
    pub fn new() -> Self {
        VbaVm {
            frames: VecDeque::new(),
            next_frame_id: 0,
            vm_state: VmState::Running,
            saved_error_frame: None,
        }
    }

    /// Push a new frame onto the stack.
    pub fn push_frame(&mut self, kind: FrameKind, list_id: usize, statements: Vec<Statement>) {
        let depth = self.frames.len();
        let frame = Frame::new(self.next_frame_id, kind, list_id, statements, depth);
        self.next_frame_id += 1;
        let _frame_id = frame.id;
        let _kind_debug = format!("{:?}", frame.kind);
        self.frames.push_back(frame);

        // eprintln!("📍 VM: pushed frame #{} {} (depth={})", frame_id, kind_debug, depth);
    }

    /// Pop the current frame (if any).
    pub fn pop_frame(&mut self) -> Option<Frame> {
        let frame = self.frames.pop_back();
        if let Some(ref _f) = frame {
            // eprintln!("📍 VM: popped frame #{} (depth now={})", f.id, self.frames.len());
        }
        frame
    }

    /// Get the current (top) frame.
    pub fn current_frame(&self) -> Option<&Frame> {
        self.frames.back()
    }

    /// Get the current frame mutably.
    pub fn current_frame_mut(&mut self) -> Option<&mut Frame> {
        self.frames.back_mut()
    }

    /// Get frame by ID (for finding Resume targets).
    pub fn frame_by_id(&self, id: usize) -> Option<&Frame> {
        self.frames.iter().find(|f| f.id == id)
    }

    /// Transition to error state.
    pub fn enter_error_state(&mut self, error_label: String, from_frame_id: usize) {
        // eprintln!("💥 VM: entering error state, label='{}', from_frame=#{}", error_label, from_frame_id);
        self.vm_state = VmState::ErrorInProgress {
            error_label,
            from_frame_id,
        };
    }

    /// Transition to resume state (handler found and about to execute Resume).
    pub fn enter_resume_state(&mut self, target_frame_id: usize, target_pc: usize) {
        // eprintln!("🔄 VM: entering resume state, target_frame=#{}, pc={}", target_frame_id, target_pc);
        self.vm_state = VmState::ResumingFromError {
            target_frame_id,
            target_pc,
        };
    }

    /// Return to running state.
    pub fn resume_running(&mut self) {
        // eprintln!("▶️ VM: resuming normal execution");
        self.vm_state = VmState::Running;
    }

    /// Check if we're in error state looking for a handler.
    pub fn is_in_error_state(&self) -> bool {
        matches!(self.vm_state, VmState::ErrorInProgress { .. })
    }

    /// Check if we're resuming from error.
    pub fn is_resuming(&self) -> bool {
        matches!(self.vm_state, VmState::ResumingFromError { .. })
    }

    /// Get error label if in error state.
    pub fn error_label(&self) -> Option<&str> {
        match &self.vm_state {
            VmState::ErrorInProgress { error_label, .. } => Some(error_label),
            _ => None,
        }
    }
    pub fn save_error_frame(&mut self, frame_id: usize) {
        if let Some(pos) = self.frames.iter().position(|f| f.id == frame_id) {
            if let Some(frame) = self.frames.remove(pos) {
                // eprintln!(
                //     "💾 VM: saving error frame #{} (pc={})",
                //     frame.id, frame.pc
                // );
                self.saved_error_frame = Some(frame);
            } else {
                // eprintln!(
                //     "⚠️ VM: save_error_frame: remove(pos={}) returned None",
                //     pos
                // );
            }
        } else {
            // eprintln!("⚠️ VM: save_error_frame: frame #{} not found", frame_id);
        }
    }

        /// Take the stashed error frame, if any.
    pub fn take_saved_error_frame(&mut self) -> Option<Frame> {
        self.saved_error_frame.take()
    }

    
    
}
pub fn execute_for_frame(
    frame: &mut Frame,
    _for_stmt: &crate::ast::ForStatement,
    ctx: &mut Context,
) -> Option<ControlFlow> {
    // Extract loop state from frame.kind
    if let FrameKind::For { counter, current_value, end_value, step } = &frame.kind {
        let counter_name = counter.clone();
        let current = *current_value;
        let end = *end_value;
        let step = *step;
        
        // Check condition
        let should_continue = if step > 0 {
            current <= end
        } else {
            current >= end
        };
        
        if !should_continue {
            return Some(ControlFlow::ExitFor);  // Loop done
        }
        
        // Set counter variable
        ctx.set_var(counter_name.clone(), crate::context::Value::Integer(current));
        
        // Execute next statement in body
        if frame.pc >= frame.statements.len() {
            // We've finished the body; step and start over
            let next_value = current + step;
            if let FrameKind::For { current_value, .. } = &mut frame.kind {
                *current_value = next_value;
            }
            frame.pc = 0;  // Restart body
            return Some(ControlFlow::Continue);  // Keep looping
        }
        
        // Normal statement execution (returned from execute_statement_in_vm)
        return None;  // Let main loop handle it
    }
    None
}
/// Execute a statement list using the VM.
/// Called from interpreter/mod.rs via run_subroutine.
pub fn run_statement_list_vm(
    stmts: &[Statement],
    ctx: &mut Context,
    list_id: usize,
) -> ControlFlow {
    let mut vm = VbaVm::new();
    vm.push_frame(FrameKind::Main, list_id, stmts.to_vec());
    // eprintln!("📋 Frame #0 statements:");
    // for (i, stmt) in stmts.iter().enumerate() {
    //     eprintln!("  [{}]: {:?}", i, stmt);
    // }

    loop {
        // 1) Check if frames left
        if vm.frames.is_empty() {
            // eprintln!("✅ VM: all frames popped, execution complete");
            return ControlFlow::Continue;
        }

        // 2) Handle error state
        if vm.is_in_error_state() {
            // eprintln!("💥 VM: error state active, searching for handler label");
        
            if let Some(label_str) = vm.error_label() {
                let label = label_str.to_string();
                let mut found = false;
                
                if let Some(loc) = &ctx.resume_location {
                    vm.save_error_frame(loc.frame_id);
                }
        
                if !vm.frames.is_empty() {
                    for i in (0..vm.frames.len()).rev() {
                        if let Some(target_pc) = find_label_in_statements(&vm.frames[i].statements, &label) {
                            // eprintln!("✅ VM: found handler at frame index {}, pc={}", i, target_pc);
                            
                            while vm.frames.len() > i + 1 {
                                vm.pop_frame();
                            }
                            
                            let handler_frame_id = vm.frames[i].id;
                            vm.frames[i].jump_to(target_pc);
                            vm.enter_resume_state(handler_frame_id, target_pc);
                            found = true;
                            break;
                        }
                    }
                }
                if !found {
                    // eprintln!("❌ VM: no handler found anywhere, exiting Sub");
                    return ControlFlow::ExitSub;
                }
            }
            continue;
        }

        // 3) Handle resume state
        if vm.is_resuming() {
            if let VmState::ResumingFromError { target_frame_id, target_pc } = vm.vm_state.clone() {
                if let Some(frame) = vm.current_frame() {
                    if frame.id == target_frame_id {
                        // eprintln!("✅ VM: resume reached target frame #{}, setting pc to {}", target_frame_id, target_pc + 1);
                        vm.current_frame_mut().unwrap().jump_to(target_pc + 1);
                        vm.resume_running();
                        continue;
                    }
                }
            }
        }

        // 4) Get current statement
        let frame = vm.current_frame().unwrap();
        let current_stmt = match frame.current_statement() {
            Some(stmt) => stmt.clone(),
            None => {
               // Frame is done
                let frame_kind = vm.current_frame().map(|f| f.kind.clone());
                vm.pop_frame();
                
                // ✅ If this was a Block (subroutine), pop scope
                if matches!(frame_kind, Some(FrameKind::Block)) {
                    ctx.pop_scope();
                    // ✅ DON'T advance parent for Block frames - already advanced when pushed
                } else {
                    // ✅ For other frame types (For, Do), DO advance parent
                    if let Some(parent) = vm.current_frame_mut() {
                        parent.advance();
                    }
                }
                continue;
            }
        };

        // eprintln!("▶️ [frame #{}] pc={} stmt={:?}", frame.id, frame.pc, current_stmt);

        // 5) Execute statement
        let flow = execute_statement_in_vm(&current_stmt, ctx, &mut vm);
        // eprintln!("  ↳ flow: {:?}", flow);
        // if ctx.err.is_some() {
        //     eprintln!("  ⚠️ ctx.err = {:?}", ctx.err);
        // }

        // 5.5) Check if an error was set during expression evaluation and we have error handling
        // This catches errors from operations like division by zero that set ctx.err but return Ok(...)
        // Skip if resume_valid is true (we're already handling this error) or if we're processing ResumeNext
        if ctx.err.is_some() && flow == ControlFlow::Continue && !ctx.resume_valid {
            if ctx.on_error_mode == crate::context::OnErrorMode::GoTo {
                if let Some(label) = ctx.on_error_label.clone() {
                    let error_frame_id = vm.current_frame().map(|f| f.id).unwrap_or(0);
                    let error_pc = vm.current_frame().map(|f| f.pc).unwrap_or(0);
                    let parent_pc = if vm.frames.len() >= 2 {
                        Some(vm.frames[vm.frames.len() - 2].pc)
                    } else {
                        None
                    };
                    ctx.resume_pc = Some(error_pc);
                    ctx.resume_valid = true;
                    ctx.resume_location = Some(crate::context::ResumeLocation {
                        frame_id: error_frame_id,
                        pc: error_pc,
                        parent_pc,
                    });
                    vm.enter_error_state(label, error_frame_id);
                    continue;
                }
            } else if ctx.on_error_mode == crate::context::OnErrorMode::ResumeNextAuto {
                // In Resume Next mode, clear the error and continue
                // Error info is preserved in ctx.err for Err object access
            }
        }

        // 6) Handle control flow
        match flow {
            ControlFlow::FramePushed => {
                if vm.frames.len() >= 2 {
                    let parent_idx = vm.frames.len() - 2;
                    vm.frames[parent_idx].advance();
                }
                continue;
            }
            
            ControlFlow::Continue => {
                if let Some(frame) = vm.current_frame_mut() {
                    frame.advance();
                }
            }

            ControlFlow::ErrorGoToLabel(label) => {
                let error_frame_id = vm.current_frame().map(|f| f.id).unwrap_or(0);
                let error_pc = ctx.resume_pc.unwrap_or(0);
                // ✅ Get parent frame's current PC
                let parent_pc = if vm.frames.len() >= 2 {
                    Some(vm.frames[vm.frames.len() - 2].pc)
                } else {
                    None
                };
                // eprintln!("⚡ VM: ErrorGoToLabel '{}' from frame #{}, pc={}", label, error_frame_id, error_pc);
                ctx.resume_location = Some(crate::context::ResumeLocation {
                    frame_id: error_frame_id,
                    pc: error_pc,
                    parent_pc,
                });
                // ✅ CRITICAL: Mark that we're handling this error
                //ctx.resume_valid = true;
                
                // ✅ DON'T clear ctx.err here - handler needs to read it
                // But we should mark that it's being handled
                
                vm.enter_error_state(label, error_frame_id);
                continue;
            }

            ControlFlow::GoToLabel(label) => {
                let is_error_goto = ctx.err.is_some()
                    && ctx.on_error_mode == crate::context::OnErrorMode::GoTo
                    && ctx.resume_valid
                    && ctx.resume_pc.is_some();
                
                if is_error_goto {
                    // eprintln!("🚨 VM: GoToLabel '{}' is error handler jump", label);
                    let error_frame_id = vm.current_frame().map(|f| f.id).unwrap_or(0);
                    let error_pc = ctx.resume_pc.unwrap_or(0);
                    let parent_pc = if vm.frames.len() >= 2 {
                        Some(vm.frames[vm.frames.len() - 2].pc)
                    } else {
                        None
                    };
                    vm.enter_error_state(label.clone(), error_frame_id);
                    ctx.resume_location = Some(crate::context::ResumeLocation {
                        frame_id: error_frame_id,
                        pc: error_pc,
                        parent_pc,
                    });
                    continue;
                }
                
                if let Some(frame) = vm.current_frame_mut() {
                    if let Some(target_pc) = find_label_in_statements(&frame.statements, &label) {
                        // eprintln!("✅ VM: label '{}' found in current frame at pc={}", label, target_pc);
                        frame.jump_to(target_pc);
                        continue;
                    }
                }
            
                let mut found = false;
                for i in (0..vm.frames.len() - 1).rev() {
                    if let Some(target_pc) = find_label_in_statements(&vm.frames[i].statements, &label) {
                        // eprintln!("✅ VM: label '{}' found in parent frame at pc={}", label, target_pc);
                        while vm.frames.len() > i + 1 {
                            vm.pop_frame();
                        }
                        vm.frames[i].jump_to(target_pc);
                        found = true;
                        break;
                    }
                }

                if !found {
                    // eprintln!("❌ VM: label '{}' not found in any frame, exiting", label);
                    return ControlFlow::GoToLabel(label);
                }
            }
            ControlFlow::ResumeNext => {
                // eprintln!("🔄 VM: ResumeNext - resume_location={:?}", ctx.resume_location);
                if let Some(loc) = &ctx.resume_location {
                    if let Some(target_idx) = vm.frames.iter().position(|f| f.id == loc.frame_id) {
                        while vm.frames.len() > target_idx + 1 {
                            vm.pop_frame();
                        }
                        if let Some(frame) = vm.current_frame_mut() {
                            frame.jump_to(loc.pc + 1);
                        }
                        // let parent_pc = if vm.frames.len() >= 2 {
                        //     Some(vm.frames[vm.frames.len() - 2].pc)
                        // } else {
                        //     None
                        // };
                        // // ✅ RESTORE PARENT PC (for live frame case)
                        // if let Some(parent_pc) = loc.parent_pc {
                        //     if vm.frames.len() >= 2 {
                        //         let parent_idx = vm.frames.len() - 2;
                        //         eprintln!("   ↳ Restoring parent frame #{} to pc={}", vm.frames[parent_idx].id, parent_pc);
                        //         vm.frames[parent_idx].pc = parent_pc;
                        //     }
                        // }

                        ctx.resume_valid = false;
                        ctx.resume_location = None;
                        ctx.err = None; // ✅ Clear the error after successful resume
                        vm.resume_running();
                        continue;
                    }
            
                    if let Some(mut frame) = vm.take_saved_error_frame() {
                        // eprintln!("✅ VM: restoring saved error frame #{} at pc+1={}", frame.id, loc.pc + 1);
                        frame.pc = loc.pc + 1;

                        // // ✅ SAVE parent_pc before pushing frame
                        // let parent_pc_to_restore = loc.parent_pc;
                        if let Some(parent_pc) = loc.parent_pc {
                            if let Some(parent) = vm.current_frame_mut() {
                                // eprintln!("   ↳ Restoring parent frame #{} to pc={}", parent.id, parent_pc);
                                parent.pc = parent_pc;
                            }
                        }

                        vm.frames.push_back(frame);

                        // // ✅ RESTORE PARENT PC (for saved frame case)
                        // if let Some(parent_pc) = parent_pc_to_restore {
                        //     if vm.frames.len() >= 2 {
                        //         let parent_idx = vm.frames.len() - 2;
                        //         eprintln!("   ↳ Restoring parent frame #{} to pc={}", vm.frames[parent_idx].id, parent_pc);
                        //         vm.frames[parent_idx].pc = parent_pc;
                        //     }
                        // }

                        ctx.resume_valid = false;
                        ctx.resume_location = None;
                        ctx.err = None; // ✅ Clear the error after successful resume
                        vm.resume_running();
                        continue;
                    }
            
                    // eprintln!("❌ VM: resume target frame not found");
                    return ControlFlow::ResumeNext;
                }
            }

            ControlFlow::ExitFor => {
                // eprintln!("🚪 VM: ExitFor");
                vm.pop_frame();
                // if let Some(parent) = vm.current_frame_mut() {
                //     parent.advance();
                // }
                continue;
            }

            ControlFlow::ExitDo => {
                // eprintln!("🚪 VM: ExitDo");
                vm.pop_frame();
                if let Some(parent) = vm.current_frame_mut() {
                    parent.advance();
                }
            }

            ControlFlow::ExitSub | ControlFlow::ExitFunction | ControlFlow::ExitProperty => {
                // eprintln!("🚪 VM: {:?}", flow);
                 // Pop the current frame (the Sub/Function being exited)
                vm.pop_frame();
                
                // Pop the scope
                ctx.pop_scope();
                
                // If there are still frames, advance the parent and continue
                if !vm.frames.is_empty() {
                    // if let Some(parent) = vm.current_frame_mut() {
                    //     parent.advance();
                    // }
                    continue;  // Continue execution in parent frame
                }
                
                // No more frames, exit completely
                return flow;
            }

            other => {
                // eprintln!("⚠️ VM: unhandled control flow {:?}, exiting", other);
                return other;
            }
        }

        // 7) Loop control logic
        if !vm.frames.is_empty() {
            let frame_kind = vm.current_frame().map(|f| f.kind.clone());
            
            match frame_kind {
                Some(FrameKind::For { counter, current_value, end_value, step }) => {
                    let body_len = vm.current_frame().unwrap().statements.len();
                    let body_complete = vm.current_frame().unwrap().pc >= body_len;

                    if body_complete {
                        let should_continue = if step > 0 {
                            current_value + step <= end_value
                        } else {
                            current_value + step >= end_value
                        };

                        if should_continue {
                            if let Some(frame) = vm.current_frame_mut() {
                                if let FrameKind::For { current_value: cv, step: s, .. } = &mut frame.kind {
                                    *cv += *s;
                                }
                                frame.pc = 0;
                            }
                            ctx.set_var(counter.clone(), crate::context::Value::Integer(current_value + step));
                        } else {
                            vm.pop_frame();
                            // if let Some(parent) = vm.current_frame_mut() {
                            //     parent.advance();
                            // }
                        }
                    }
                }

                Some(FrameKind::Do { ref statement, first_iteration: _ }) => {
                    let body_len = vm.current_frame().unwrap().statements.len();
                    let at_start = vm.current_frame().unwrap().pc == 0;
                    let body_complete = vm.current_frame().unwrap().pc >= body_len;

                    // Pre-test: check condition at start (except first iteration)
                    if !statement.test_at_end && at_start  {
                        // eprintln!("  ↳ Do loop (pre-test): checking condition at start");
                        match should_do_loop_continue(statement, ctx) {
                            Ok(true) => {
                                // eprintln!("     Condition true, continuing loop");
                                // Mark first iteration as done
                                if let Some(frame) = vm.current_frame_mut() {
                                    if let FrameKind::Do { first_iteration: fi, .. } = &mut frame.kind {
                                        *fi = false;
                                    }
                                }
                            }
                            Ok(false) => {
                                // eprintln!("     Condition false, exiting loop");
                                vm.pop_frame();
                                if let Some(parent) = vm.current_frame_mut() {
                                    parent.advance();
                                }
                                continue;
                            }
                            Err(e) => {
                                // eprintln!("     Error evaluating condition: {}", e);
                                ctx.err = Some(crate::context::ErrObject {
                                    number: 13,
                                    description: e,
                                    source: "Interpreter".into(),
                                });
                            }
                        }
                    }

                    // Body complete: check post-test condition or restart
                    if body_complete {
                        // eprintln!("  ↳ Do loop body complete");
                        
                        if statement.test_at_end {
                            // Post-test: check condition after body
                            match should_do_loop_continue(statement, ctx) {
                                Ok(true) => {
                                    // eprintln!("     Post-test condition true, restarting loop");
                                    if let Some(frame) = vm.current_frame_mut() {
                                        if let FrameKind::Do { first_iteration: fi, .. } = &mut frame.kind {
                                            *fi = false;
                                        }
                                        frame.pc = 0;
                                    }
                                }
                                Ok(false) => {
                                    // eprintln!("     Post-test condition false, exiting loop");
                                    vm.pop_frame();
                                    if let Some(parent) = vm.current_frame_mut() {
                                        parent.advance();
                                    }
                                }
                                Err(e) => {
                                    // eprintln!("     Error evaluating condition: {}", e);
                                    ctx.err = Some(crate::context::ErrObject {
                                        number: 13,
                                        description: e,
                                        source: "Interpreter".into(),
                                    });
                                }
                            }
                        } else {
                            // ✅ FIX: Pre-test loop - check condition BEFORE restarting
                            // eprintln!("     Pre-test loop body complete, checking condition before restart");
                            match should_do_loop_continue(statement, ctx) {
                                Ok(true) => {
                                    // eprintln!("     Condition true, restarting loop");
                                    if let Some(frame) = vm.current_frame_mut() {
                                        frame.pc = 0;
                                    }
                                }
                                Ok(false) => {
                                    // eprintln!("     Condition false, exiting loop");
                                    vm.pop_frame();
                                    if let Some(parent) = vm.current_frame_mut() {
                                        parent.advance();
                                    }
                                }
                                Err(e) => {
                                    // eprintln!("     Error evaluating condition: {}", e);
                                    ctx.err = Some(crate::context::ErrObject {
                                        number: 13,
                                        description: e,
                                        source: "Interpreter".into(),
                                    });
                                }
                            }
                        }
                    }
                }

                _ => {}
            }
        }
    }
}

/// Execute a single statement in the VM context.
/// This is called from run_statement_list_vm and dispatches to statement handlers.
fn execute_statement_in_vm(
    stmt: &Statement,
    ctx: &mut Context,
    vm: &mut VbaVm,
) -> ControlFlow {
    use crate::ast::Statement;

    match stmt {
        Statement::For(for_stmt) => {
            // eprintln!("📍 execute_statement_in_vm: FOR arm");
            crate::vm::runtime::handle_for_statement(for_stmt, ctx, vm)
        }

        Statement::DoWhile(do_stmt) => {
            handle_do_statement(do_stmt, ctx, vm)
        }

        Statement::If { condition, then_branch, else_if, else_branch } => {
            handle_if_statement(condition, then_branch, else_if, else_branch, ctx, vm)
        }
        Statement::Call { function, args } => {
            // eprintln!("📍 execute_statement_in_vm: CALL arm");
            handle_call_statement(function, args, ctx, vm)
        }

        Statement::With { object, body } => {
            handle_with_statement(object, body, ctx, vm)
        }

        // For all other statements, delegate to existing execute_statement
        _ =>{
            // eprintln!("📍 execute_statement_in_vm: delegating to interpreter");
            let pc = vm.current_frame().map(|f| f.pc).unwrap_or(0);
            crate::interpreter::execute_statement(stmt, ctx, pc)
        }
    }
}

// ============================================================================
// Helper Functions
// ============================================================================

/// Create a Do/While loop frame
fn handle_do_statement(
    do_stmt: &crate::ast::DoWhileStatement,
    ctx: &mut Context,
    vm: &mut VbaVm,
) -> ControlFlow {
    // eprintln!("📍 VM handle_do_statement: entering");
     // ✅ For pre-test loops, check condition BEFORE pushing frame
     if !do_stmt.test_at_end {
        // eprintln!("   Checking pre-test condition before entering loop");
        match should_do_loop_continue(do_stmt, ctx) {
            Ok(false) => {
                // eprintln!("   Condition false, not entering loop");
                return ControlFlow::Continue;  // Don't push frame at all!
            }
            Ok(true) => {
                // eprintln!("   Condition true, entering loop");
            }
            Err(e) => {
                // eprintln!("   Error evaluating condition: {}", e);
                ctx.err = Some(crate::context::ErrObject {
                    number: 13,
                    description: e,
                    source: "Interpreter".into(),
                });
                return ControlFlow::Continue;
            }
        }
    }
    
    vm.push_frame(
        FrameKind::Do {
            statement: do_stmt.clone(),
            first_iteration: true,
        },
        vm.next_frame_id,
        do_stmt.body.clone(),
    );
    
    ControlFlow::FramePushed
}

fn should_do_loop_continue(
    do_stmt: &crate::ast::DoWhileStatement,
    ctx: &mut Context,
) -> Result<bool, String> {
    use crate::ast::DoWhileConditionType;
    
    match &do_stmt.condition {
        Some(cond_expr) => {
            match crate::interpreter::evaluate_expression(cond_expr, ctx) {
                Ok(val) => {
                    let truthy = is_truthy(&val);
                    match do_stmt.condition_type {
                        DoWhileConditionType::While => Ok(truthy),
                        DoWhileConditionType::Until => Ok(!truthy),
                        DoWhileConditionType::Infinite => Ok(true),
                    }
                }
                Err(e) => Err(e.to_string()),
            }
        }
        None => Ok(true),
    }
}

/// Handle With block execution
fn handle_with_statement(
    object: &crate::ast::Expression,
    body: &[Statement],
    ctx: &mut Context,
    vm: &mut VbaVm,
) -> ControlFlow {
    // Evaluate the With object expression
    match crate::interpreter::evaluate_expression(object, ctx) {
        Ok(obj_value) => {
            // Push the object onto the With stack
            ctx.with_stack.push(obj_value);
            
            // Push a new frame for the With block body
            let list_id = vm.next_frame_id;
            vm.push_frame(FrameKind::With, list_id, body.to_vec());
            
            // The With object will be popped when the frame completes
            // Return FramePushed so parent advances but new frame doesn't skip first statement
            ControlFlow::FramePushed
        }
        Err(e) => {
            let pc = vm.current_frame().map(|f| f.pc).unwrap_or(0);
            ctx.err = Some(crate::context::ErrObject {
                number: 91,
                description: format!("With object evaluation failed: {}", e),
                source: "VM".into(),
            });
            // Simple error handling - just continue
            ControlFlow::Continue
        }
    }
}

/// Create an If block frame
fn handle_if_statement(
    condition: &crate::ast::Expression,
    then_branch: &[Statement],
    else_if: &[(crate::ast::Expression, Vec<Statement>)],
    else_branch: &[Statement],
    ctx: &mut Context,
    vm: &mut VbaVm,
) -> ControlFlow {
    // eprintln!("📍 If statement: evaluating condition");
    // eprintln!("   Condition: {:?}", condition);
    // Evaluate the main condition
    let cond_result = crate::interpreter::evaluate_expression(condition, ctx);
    // eprintln!("   Result: {:?}", cond_result);
    match cond_result {
        Ok(val) => {
            let is_true = is_truthy(&val);
            // eprintln!("   Condition result: {:?}, truthy: {}", val, is_true);
            
            if is_true {
                // eprintln!("   Executing then_branch with {} statements", then_branch.len());
                // Execute then branch
                for stmt in then_branch {
                    let flow = execute_statement_in_vm(stmt, ctx, vm);
                    if flow != ControlFlow::Continue {
                        return flow;
                    }
                }
                return ControlFlow::Continue;
            }
            
            // Check else-if conditions
            for (elseif_cond, elseif_stmts) in else_if {
                let elseif_result = crate::interpreter::evaluate_expression(elseif_cond, ctx);
                if let Ok(elseif_val) = elseif_result {
                    if is_truthy(&elseif_val) {
                        // eprintln!("   Executing else-if branch");
                        for stmt in elseif_stmts {
                            let flow = execute_statement_in_vm(stmt, ctx, vm);
                            if flow != ControlFlow::Continue {
                                return flow;
                            }
                        }
                        return ControlFlow::Continue;
                    }
                }
            }
            
            // Execute else branch if present
            if !else_branch.is_empty() {
                // eprintln!("   Executing else_branch with {} statements", else_branch.len());
                for stmt in else_branch {
                    let flow = execute_statement_in_vm(stmt, ctx, vm);
                    if flow != ControlFlow::Continue {
                        return flow;
                    }
                }
            }
            
            ControlFlow::Continue
        }
        Err(e) => {
            // eprintln!("   Error evaluating condition: {}", e);
            // Set error in context
            ctx.err = Some(crate::context::ErrObject {
                number: 13,
                description: e.to_string(),
                source: "Interpreter".into(),
            });
            ControlFlow::Continue
        }
    }
}

/// Helper function to check if a value is truthy
fn is_truthy(v: &crate::context::Value) -> bool {
    use crate::context::Value;
    match v {
        Value::Boolean(b) => *b,
        Value::Integer(i) => *i != 0,
        Value::Long(i) => *i != 0,
        Value::LongLong(i) => *i != 0,
        Value::Object(None) => false,
        Value::Object(Some(inner)) => is_truthy(inner),
        Value::Byte(b) => *b != 0,
        Value::Currency(c) => *c != 0.0,
        Value::Date(_) => true,
        Value::DateTime(_) => true,
        Value::Time(_) => true,
        Value::Double(f) => *f != 0.0,
        Value::Decimal(f) => *f != 0.0,
        Value::Single(f) => *f != 0.0,
        Value::String(s) => !s.is_empty(),
        Value::UserType { .. } => true,
        Value::Empty => false,
        Value::Null => false,
        Value::Error(_) => false,  // Error values are falsy
    }
}


/// Find a label in a frame's statements.
fn find_label_in_frame(frame: &Frame, label: &str) -> Option<usize> {
    find_label_in_statements(&frame.statements, label)
}

/// Find a label in a statement list.
fn find_label_in_statements(stmts: &[Statement], label: &str) -> Option<usize> {
    let target = label.to_ascii_lowercase();

    // 1) Exact (case-insensitive) match first – the correct / future path
    if let Some(idx) = stmts.iter().enumerate().find_map(|(idx, stmt)| {
        if let Statement::Label(name) = stmt {
            if name.to_ascii_lowercase() == target {
                return Some(idx);
            }
        }
        None
    }) {
        return Some(idx);
    }

    // 2) Fallback: suffix match to work around labels like "Point" for "ExitPoint"
    let mut fallback_idx: Option<usize> = None;

    for (idx, stmt) in stmts.iter().enumerate() {
        if let Statement::Label(name) = stmt {
            let name_lower = name.to_ascii_lowercase();
            if target.ends_with(&name_lower) {
                // If we already have a fallback, it's ambiguous → bail
                if fallback_idx.is_some() {
                    // eprintln!("⚠️ Ambiguous label match: target='{}', candidate='{}'", label, name);
                    return None;
                }
                fallback_idx = Some(idx);
            }
        }
    }

    fallback_idx
}

// In vm/runtime.rs, add a helper that is called from execute_statement_in_vm:

pub fn handle_for_statement(
    for_stmt: &crate::ast::ForStatement,
    ctx: &mut Context,
    vm: &mut VbaVm,
) -> ControlFlow {
    // eprintln!("📍 VM handle_for_statement: entering");
    use crate::context::Value;

    // Evaluate start, end, step using the existing expression evaluator
    let start = crate::interpreter::evaluate_expression(&for_stmt.start, ctx).ok();
    let end = crate::interpreter::evaluate_expression(&for_stmt.end, ctx).ok();
    let step_expr: Value = for_stmt
        .step
        .as_ref()
        .and_then(|e| crate::interpreter::evaluate_expression(e, ctx).ok())
        .unwrap_or(Value::Integer(1));

    // Coerce to integers using the shared helper; fall back to simple defaults
    let start_int = start
        .as_ref()
        .and_then(|v| crate::interpreter::value_to_integer(v).ok())
        .unwrap_or(0);
    let end_int = end
        .as_ref()
        .and_then(|v| crate::interpreter::value_to_integer(v).ok())
        .unwrap_or(0);
    let step_int = crate::interpreter::value_to_integer(&step_expr).unwrap_or(1);

    // Push For frame
    vm.push_frame(
        FrameKind::For {
            counter: for_stmt.counter.clone(),
            current_value: start_int,
            end_value: end_int,
            step: step_int,
        },
        /* list_id */ vm.next_frame_id, // or better: list_id passed into run_statement_list_vm
        for_stmt.body.clone(),
    );
    ctx.set_var(for_stmt.counter.clone(), Value::Integer(start_int));

    // eprintln!("📍 VM handle_for_statement: returning FramePushed");
    ControlFlow::FramePushed
}
fn handle_call_statement(
    function: &str,
    args: &[Expression],
    ctx: &mut Context,
    vm: &mut VbaVm,
) -> ControlFlow {
    // Handle builtins
    // eprintln!("📍 Call statement: ");
    
    if handle_builtin_call_bool(function, args, ctx) {
        return ControlFlow::Continue;
    }

    // Get sub definition
    let (params, body) = match ctx.subs.get(function).cloned() {
        Some(pb) => pb,
        None => return ControlFlow::Continue,
    };

    // Evaluate arguments
    let mut arg_vals = Vec::new();
    for a in args {
        match crate::interpreter::evaluate_expression(a, ctx) {
            Ok(v) => arg_vals.push(v),
            Err(_) => return ControlFlow::Continue,
        }
    }

    // Push scope
    ctx.push_scope(function.to_string(), ScopeKind::Subroutine);
    
    // Bind parameters
    for (param, val) in params.iter().zip(arg_vals) {
        ctx.declare_variable(&param.name);  // Use param.name for Parameter struct
        ctx.declare_local(param.name.clone(), val);
    }

    // ✅ Push VM frame for subroutine
    vm.push_frame(FrameKind::Block, vm.next_frame_id, body);
    // eprintln!("📍 VM handle_call_statement: returning FramePushed");
    ControlFlow::FramePushed
}
