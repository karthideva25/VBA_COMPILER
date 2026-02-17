mod constants;
pub mod functions;

pub(crate) use constants::resolve_builtin_identifier;
pub(crate) use functions::handle_builtin_call_bool;
