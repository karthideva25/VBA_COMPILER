// src/interpreter/mod.rs
mod expressions;
mod statements;
mod operations;
mod coerce;

pub mod builtins;
// pub mod host;

pub(crate) use expressions::evaluate_expression;
pub use statements::execute_statement_list;
pub use crate::vm::run_statement_list_vm;  // ← ADD THIS

// Re-export core control-flow and helpers so other modules (like `vm`) can use them
pub use self::statements::ControlFlow;
pub(crate) use self::statements::execute_statement;
pub use self::statements::value_to_integer;

use crate::ast::{Program, Statement};
use crate::context::Context;
use anyhow::Result;

pub fn execute_ast(program: &Program, ctx: &mut Context) -> Result<()> {
    for stmt in &program.statements {
        if let Statement::Subroutine { name, params, body } = stmt {
            ctx.subs.insert(name.clone(), (params.clone(), body.clone()));
        }
    }
    Ok(())
}

/// Updated to use the VM
pub fn run_subroutine(ctx: &mut Context, name: &str) {
    let body: Vec<Statement> = match ctx.subs.get(name) {
        Some((_params, body)) => body.clone(),
        None => {
            eprintln!("Subroutine '{}' not found", name);
            return;
        }
    };

    println!("Entering Sub {}", name);

    // ← USE THE VM HERE
    let flow = run_statement_list_vm(&body, ctx, 0);

    println!("Leaving Sub {}", name);

    match flow {
        ControlFlow::Continue
        | ControlFlow::ExitSub
        | ControlFlow::ResumeNext => {
            // Normal termination
        }
        other => {
            eprintln!("Subroutine '{}' finished with control flow: {:?}", name, other);
        }
    }
}