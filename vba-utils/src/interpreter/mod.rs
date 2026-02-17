mod expressions;
mod statements;
mod operations;
mod coerce;

pub mod builtins;
pub mod host;

pub(crate) use expressions::evaluate_expression;
pub use statements::execute_statement_list;

use crate::ast::Program;
use crate::context::Context;
use anyhow::Result;

pub fn execute_ast(program: &Program, ctx: &mut Context) -> Result<()> {
    // drive the whole module with the list executor
    let _ = execute_statement_list(&program.statements, ctx);
    Ok(())
}