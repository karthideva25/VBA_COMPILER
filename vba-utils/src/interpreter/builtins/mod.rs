mod constants;
pub mod functions;

// Category-specific function modules
mod common;
mod strings;
mod datetime;
mod math;
mod conversion;
mod information;
mod interaction;
mod financial;
mod errobj;

pub(crate) use constants::resolve_builtin_identifier;
pub(crate) use functions::handle_builtin_call_bool;
pub(crate) use errobj::handle_err_method;
pub(crate) use errobj::handle_err_function;
