//! The wall clock the engine reads for `NOW()`, `TODAY()`, the seed of `RAND()` and the build time.
//!
//! On a machine with an operating system it is the system clock. On `wasm32-unknown-unknown`
//! (a browser or Node) `std::time::SystemTime::now()` panics, so the host installs a clock once
//! with [`set_wall_clock`] (the WASM package installs `Date.now()`); until then the clock reads
//! the Unix epoch, so that a missing clock gives a fixed answer and never a panic.

use std::sync::OnceLock;
use std::time::Duration;

static WALL_CLOCK: OnceLock<fn() -> f64> = OnceLock::new();

/// Install the function that gives the seconds since the Unix epoch (fractions allowed). It can be
/// set once; a second call is ignored.
pub fn set_wall_clock(clock: fn() -> f64) {
    let _ = WALL_CLOCK.set(clock);
}

/// The time since the Unix epoch.
#[must_use]
pub fn unix_now() -> Duration {
    if let Some(clock) = WALL_CLOCK.get() {
        let seconds = clock();
        return if seconds.is_finite() && seconds > 0.0 {
            Duration::from_secs_f64(seconds)
        } else {
            Duration::ZERO
        };
    }
    #[cfg(all(target_arch = "wasm32", target_os = "unknown"))]
    {
        Duration::ZERO
    }
    #[cfg(not(all(target_arch = "wasm32", target_os = "unknown")))]
    {
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or(Duration::ZERO)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_clock_reads_something_after_2020() {
        assert!(unix_now().as_secs() > 1_577_836_800);
    }
}
