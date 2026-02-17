//! Runtime Configuration for VBA Interpreter
//!
//! This module contains session-level configuration that is passed from the
//! application layer (web server, CLI, native app) to the interpreter.
//!
//! ## Architecture
//!
//! ```text
//! Application Layer (Web/CLI/Native)
//!         │
//!         │ Creates RuntimeConfig with:
//!         │   • User's timezone
//!         │   • User's locale
//!         │   • Workbook handle
//!         │   • User permissions
//!         ▼
//! ┌─────────────────────────────────────┐
//! │  Context::with_config(config)       │
//! │                                     │
//! │  Interpreter uses config for:       │
//! │    • Now(), Date(), Time()          │
//! │    • Format() locale settings       │
//! │    • Workbook/Range access          │
//! └─────────────────────────────────────┘
//! ```
//!
//! ## Usage
//!
//! ```rust,ignore
//! use vba_utils::{Context, RuntimeConfig};
//!
//! // In your web handler / API endpoint:
//! let config = RuntimeConfig::builder()
//!     .timezone("Asia/Kolkata")
//!     .locale("en-IN")
//!     .workbook_id("wb-12345")
//!     .user_id("user-67890")
//!     .build();
//!
//! let mut ctx = Context::with_config(config);
//! // Now execute VBA code...
//! ```

use chrono_tz::Tz;
use std::str::FromStr;

/// Runtime configuration passed from application layer to interpreter.
/// 
/// This struct contains all session-level metadata needed during VBA execution.
/// It is immutable once created - use the builder pattern to construct it.
#[derive(Debug, Clone)]
pub struct RuntimeConfig {
    /// User's timezone for Now(), Date(), Time() functions
    /// Examples: "Asia/Kolkata", "America/New_York", "Europe/London", "UTC"
    pub timezone: Tz,
    
    /// User's locale for formatting (future use)
    /// Examples: "en-US", "en-IN", "de-DE"
    pub locale: String,
    
    /// Active workbook identifier (passed to excel-host)
    pub workbook_id: Option<String>,
    
    /// Current user identifier (for audit/permissions)
    pub user_id: Option<String>,
    
    /// First day of week for date functions (1=Sunday, 2=Monday, etc.)
    /// VBA default is Sunday (1)
    pub first_day_of_week: u8,
    
    /// First week of year for date functions
    /// 1 = Week containing Jan 1 (VBA default)
    /// 2 = First week with at least 4 days
    /// 3 = First full week
    pub first_week_of_year: u8,
}

impl Default for RuntimeConfig {
    fn default() -> Self {
        Self {
            timezone: Tz::UTC,
            locale: "en-US".to_string(),
            workbook_id: None,
            user_id: None,
            first_day_of_week: 1,  // Sunday
            first_week_of_year: 1, // Week containing Jan 1
        }
    }
}

impl RuntimeConfig {
    /// Create a new RuntimeConfig with defaults (UTC timezone)
    pub fn new() -> Self {
        Self::default()
    }
    
    /// Create a builder for fluent configuration
    pub fn builder() -> RuntimeConfigBuilder {
        RuntimeConfigBuilder::new()
    }
    
    /// Quick constructor with just timezone
    pub fn with_timezone(tz_name: &str) -> Result<Self, String> {
        let tz = Tz::from_str(tz_name)
            .map_err(|_| format!("Invalid timezone: {}", tz_name))?;
        Ok(Self {
            timezone: tz,
            ..Default::default()
        })
    }
    
    /// Get the timezone name as a string
    pub fn timezone_name(&self) -> &str {
        self.timezone.name()
    }
}

/// Builder for RuntimeConfig
#[derive(Debug, Default)]
pub struct RuntimeConfigBuilder {
    timezone: Option<Tz>,
    locale: Option<String>,
    workbook_id: Option<String>,
    user_id: Option<String>,
    first_day_of_week: Option<u8>,
    first_week_of_year: Option<u8>,
}

impl RuntimeConfigBuilder {
    pub fn new() -> Self {
        Self::default()
    }
    
    /// Set the timezone by name (e.g., "Asia/Kolkata", "America/New_York")
    pub fn timezone(mut self, tz_name: &str) -> Self {
        if let Ok(tz) = Tz::from_str(tz_name) {
            self.timezone = Some(tz);
        } else {
            eprintln!("Warning: Invalid timezone '{}', using UTC", tz_name);
            self.timezone = Some(Tz::UTC);
        }
        self
    }
    
    /// Set the locale (e.g., "en-US", "en-IN")
    pub fn locale(mut self, locale: &str) -> Self {
        self.locale = Some(locale.to_string());
        self
    }
    
    /// Set the workbook ID
    pub fn workbook_id(mut self, id: &str) -> Self {
        self.workbook_id = Some(id.to_string());
        self
    }
    
    /// Set the user ID
    pub fn user_id(mut self, id: &str) -> Self {
        self.user_id = Some(id.to_string());
        self
    }
    
    /// Set the first day of week (1=Sunday, 2=Monday, ..., 7=Saturday)
    pub fn first_day_of_week(mut self, day: u8) -> Self {
        self.first_day_of_week = Some(day.clamp(1, 7));
        self
    }
    
    /// Set the first week of year (1, 2, or 3)
    pub fn first_week_of_year(mut self, week: u8) -> Self {
        self.first_week_of_year = Some(week.clamp(1, 3));
        self
    }
    
    /// Build the RuntimeConfig
    pub fn build(self) -> RuntimeConfig {
        RuntimeConfig {
            timezone: self.timezone.unwrap_or(Tz::UTC),
            locale: self.locale.unwrap_or_else(|| "en-US".to_string()),
            workbook_id: self.workbook_id,
            user_id: self.user_id,
            first_day_of_week: self.first_day_of_week.unwrap_or(1),
            first_week_of_year: self.first_week_of_year.unwrap_or(1),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    
    #[test]
    fn test_default_config() {
        let config = RuntimeConfig::default();
        assert_eq!(config.timezone, Tz::UTC);
        assert_eq!(config.locale, "en-US");
    }
    
    #[test]
    fn test_builder() {
        let config = RuntimeConfig::builder()
            .timezone("Asia/Kolkata")
            .locale("en-IN")
            .workbook_id("test-wb")
            .user_id("user-123")
            .build();
        
        assert_eq!(config.timezone_name(), "Asia/Kolkata");
        assert_eq!(config.locale, "en-IN");
        assert_eq!(config.workbook_id, Some("test-wb".to_string()));
        assert_eq!(config.user_id, Some("user-123".to_string()));
    }
    
    #[test]
    fn test_invalid_timezone_falls_back_to_utc() {
        let config = RuntimeConfig::builder()
            .timezone("Invalid/Timezone")
            .build();
        
        assert_eq!(config.timezone, Tz::UTC);
    }
}
