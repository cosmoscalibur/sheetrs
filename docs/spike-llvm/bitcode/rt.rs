#![no_std]
#[unsafe(no_mangle)] pub extern "C" fn rt_round2(x: f64) -> f64 {
    // pure numeric leaf of the runtime: half away from zero at 2 decimals
    let t = x * 100.0; let r = if t >= 0.0 { (t + 0.5) as i64 as f64 } else { (t - 0.5) as i64 as f64 }; r / 100.0
}
#[panic_handler] fn p(_: &core::panic::PanicInfo) -> ! { loop {} }
