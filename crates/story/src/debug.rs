// Re-export log macros
pub use log::{debug, error, info, warn, trace};

// For backward compatibility
#[macro_export]
macro_rules! debug {
    ($($arg:tt)*) => {
        log::debug!($($arg)*)
    };
} 