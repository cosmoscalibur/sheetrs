mod rd;
use inkwell::context::Context;
use inkwell::passes::PassBuilderOptions;
use inkwell::targets::*;
use inkwell::{FloatPredicate, OptimizationLevel, AddressSpace};
use std::time::Instant;

extern "C" fn rt_round2(x: f64) -> f64 { let t = x * 100.0; let r = if t >= 0.0 { (t + 0.5) as i64 as f64 } else { (t - 0.5) as i64 as f64 }; r / 100.0 }
extern "C" fn round2(x: f64) -> f64 { rd::round_dec(x, 2, rd::Mode::Nearest) }

type Kern = unsafe extern "C" fn(f64, i64, *const f64, *const f64, *const f64, *mut f64, *mut f64, *mut f64, *mut f64);

fn data(n: usize) -> (Vec<f64>, Vec<f64>, Vec<f64>) {
    let mut s = 12345u64; let mut rnd = || { s ^= s << 13; s ^= s >> 7; s ^= s << 17; s };
    let a = (0..n).map(|_| (rnd() % 1000 + 1) as f64).collect();
    let b = (0..n).map(|_| ((rnd() % 49900 + 100) as f64) / 100.0).collect();
    let c = (0..n).map(|_| ((rnd() % 15000 + 5000) as f64) / 10000.0).collect();
    (a, b, c)
}

fn reference(rate: f64, a: &[f64], b: &[f64], c: &[f64], exact: bool) -> (Vec<f64>, Vec<f64>, Vec<f64>, Vec<f64>) {
    let n = a.len(); let (mut d, mut e, mut g, mut h) = (vec![0.0; n], vec![0.0; n], vec![0.0; n], vec![0.0; n]);
    let mut run = 0.0;
    for i in 0..n {
        let x = a[i] * b[i]; let a1 = if x > 50000.0 { x * (1.0 - rate) } else { x };
        let y = a1 * c[i];
        d[i] = if exact { round2(y) } else { (y * 100.0).round() / 100.0 };
        run += d[i]; e[i] = run;
        let q = d[i] / a[i]; g[i] = if q.is_finite() { q } else { 0.0 };
        h[i] = d[i].max(g[i]) - d[i].min(g[i]);
    }
    (d, e, g, h)
}

fn llvm_kernel<'c>(ctx: &'c Context, exact: bool) -> inkwell::module::Module<'c> { llvm_kernel_named(ctx, exact, "kernel") }
fn llvm_kernel_named<'c>(ctx: &'c Context, exact: bool, name: &str) -> inkwell::module::Module<'c> {
    let m = ctx.create_module("k"); let b = ctx.create_builder();
    let f64t = ctx.f64_type(); let i64t = ctx.i64_type(); let ptr = ctx.ptr_type(AddressSpace::default());
    let r2t = f64t.fn_type(&[f64t.into()], false);
    let r2 = m.add_function(if std::env::var("RT").is_ok() { "rt_round2" } else { "round2" }, r2t, None);
    let ft = ctx.void_type().fn_type(&[f64t.into(), i64t.into(), ptr.into(), ptr.into(), ptr.into(), ptr.into(), ptr.into(), ptr.into(), ptr.into()], false);
    let f = m.add_function(name, ft, None);
    let entry = ctx.append_basic_block(f, "entry"); let lp = ctx.append_basic_block(f, "loop"); let done = ctx.append_basic_block(f, "done");
    b.position_at_end(entry);
    b.build_unconditional_branch(lp).unwrap();
    b.position_at_end(lp);
    let i = b.build_phi(i64t, "i").unwrap(); let run = b.build_phi(f64t, "run").unwrap();
    i.add_incoming(&[(&i64t.const_zero(), entry)]); run.add_incoming(&[(&f64t.const_zero(), entry)]);
    let iv = i.as_basic_value().into_int_value(); let rv = run.as_basic_value().into_float_value();
    let p = |k: u32| f.get_nth_param(k).unwrap().into_pointer_value();
    let at = |k: u32, name: &str| unsafe { b.build_gep(f64t, p(k), &[iv], name).unwrap() };
    let ld = |k: u32, name: &str| b.build_load(f64t, at(k, name), name).unwrap().into_float_value();
    let rate = f.get_nth_param(0).unwrap().into_float_value();
    let a = ld(2, "a"); let b0 = ld(3, "b"); let c = ld(4, "c");
    let x = b.build_float_mul(a, b0, "x").unwrap();
    let one = f64t.const_float(1.0);
    let disc = b.build_float_mul(x, b.build_float_sub(one, rate, "om").unwrap(), "disc").unwrap();
    let gt = b.build_float_compare(FloatPredicate::OGT, x, f64t.const_float(50000.0), "gt").unwrap();
    let a1 = b.build_select(gt, disc, x, "a1").unwrap().into_float_value();
    let y = b.build_float_mul(a1, c, "y").unwrap();
    let d = if exact {
        b.build_call(r2, &[y.into()], "d").unwrap().try_as_basic_value().unwrap_basic().into_float_value()
    } else {
        let hundred = f64t.const_float(100.0);
        let rnd = inkwell::intrinsics::Intrinsic::find("llvm.round").unwrap().get_declaration(&m, &[f64t.into()]).unwrap();
        let t = b.build_float_mul(y, hundred, "t").unwrap();
        let r = b.build_call(rnd, &[t.into()], "r").unwrap().try_as_basic_value().unwrap_basic().into_float_value();
        b.build_float_div(r, hundred, "d").unwrap()
    };
    let run2 = b.build_float_add(rv, d, "run2").unwrap();
    let q = b.build_float_div(d, a, "q").unwrap();
    let fabs = inkwell::intrinsics::Intrinsic::find("llvm.fabs").unwrap().get_declaration(&m, &[f64t.into()]).unwrap();
    let aq = b.build_call(fabs, &[q.into()], "aq").unwrap().try_as_basic_value().unwrap_basic().into_float_value();
    let fin = b.build_float_compare(FloatPredicate::OLT, aq, f64t.const_float(f64::INFINITY), "fin").unwrap();
    let g = b.build_select(fin, q, f64t.const_zero(), "g").unwrap().into_float_value();
    let dgt = b.build_float_compare(FloatPredicate::OGT, d, g, "dgt").unwrap();
    let mx = b.build_select(dgt, d, g, "mx").unwrap().into_float_value();
    let mn = b.build_select(dgt, g, d, "mn").unwrap().into_float_value();
    let h = b.build_float_sub(mx, mn, "h").unwrap();
    b.build_store(at(5, "pd"), d).unwrap(); b.build_store(at(6, "pe"), run2).unwrap();
    b.build_store(at(7, "pg"), g).unwrap(); b.build_store(at(8, "ph"), h).unwrap();
    let inext = b.build_int_add(iv, i64t.const_int(1, false), "inext").unwrap();
    let cont = b.build_int_compare(inkwell::IntPredicate::ULT, inext, f.get_nth_param(1).unwrap().into_int_value(), "cont").unwrap();
    i.add_incoming(&[(&inext, lp)]); run.add_incoming(&[(&run2, lp)]);
    b.build_conditional_branch(cont, lp, done).unwrap();
    b.position_at_end(done); b.build_return(None).unwrap();
    m
}

fn tm(level: OptimizationLevel) -> TargetMachine {
    Target::initialize_native(&InitializationConfig::default()).unwrap();
    let triple = TargetMachine::get_default_triple(); let t = Target::from_triple(&triple).unwrap();
    t.create_target_machine(&triple, &TargetMachine::get_host_cpu_name().to_string(), &TargetMachine::get_host_cpu_features().to_string(), level, RelocMode::Default, CodeModel::JITDefault).unwrap()
}

fn kernel_test() {
    let n = 300_000; let (a, b, c) = data(n);
    for exact in [false, true] {
        let rate = 0.05;
        let want = reference(rate, &a, &b, &c, exact);
        for (name, pipeline, lvl) in [("O0", "default<O0>", OptimizationLevel::None), ("O2", "default<O2>", OptimizationLevel::Default), ("O3", "default<O3>", OptimizationLevel::Aggressive)] {
            let ctx = Context::create();
            let t0 = Instant::now();
            let m = llvm_kernel(&ctx, exact);
            let t_ir = t0.elapsed();
            let machine = tm(lvl);
            let t1 = Instant::now();
            m.run_passes(pipeline, &machine, PassBuilderOptions::create()).unwrap();
            let t_opt = t1.elapsed();
            let t2 = Instant::now();
            let ee = m.create_jit_execution_engine(lvl).unwrap();
            if let Some(r) = m.get_function("round2") { ee.add_global_mapping(&r, round2 as usize); }
            let f: Kern = unsafe { std::mem::transmute(ee.get_function_address("kernel").unwrap()) };
            let t_cg = t2.elapsed();
            let (mut d, mut e, mut g, mut h) = (vec![0.0; n], vec![0.0; n], vec![0.0; n], vec![0.0; n]);
            let mut best = f64::MAX;
            for _ in 0..7 { let t = Instant::now(); unsafe { f(rate, n as i64, a.as_ptr(), b.as_ptr(), c.as_ptr(), d.as_mut_ptr(), e.as_mut_ptr(), g.as_mut_ptr(), h.as_mut_ptr()) }; best = best.min(t.elapsed().as_secs_f64()); }
            let same = d == want.0 && e == want.1 && g == want.2 && h == want.3;
            println!("exact_round={exact:5} {name}: build IR {:?}  passes {:?}  codegen {:?}  run best {:.2} ms  identical_to_rust={same}", t_ir, t_opt, t_cg, best * 1e3);
        }
        let mut best = f64::MAX;
        for _ in 0..7 { let t = Instant::now(); let r = reference(0.05, &a, &b, &c, exact); best = best.min(t.elapsed().as_secs_f64()); std::hint::black_box(r); }
        println!("exact_round={exact:5} rustc -O reference: run best {:.2} ms", best * 1e3);
    }
}

// ---- compile latency for N distinct small formulas ----
#[derive(Clone)] enum E { V(usize), K(f64), B(u8, Box<E>, Box<E>), S(Box<E>, Box<E>, Box<E>, Box<E>) }
fn gen(r: &mut u64, d: u32) -> E {
    let mut rnd = || { *r ^= *r << 13; *r ^= *r >> 7; *r ^= *r << 17; *r };
    if d == 0 || rnd() % 5 == 0 { return if rnd() % 3 == 0 { E::K((rnd() % 1000) as f64 / 10.0) } else { E::V((rnd() % 3) as usize) }; }
    if rnd() % 6 == 0 { return E::S(Box::new(gen(r, d - 1)), Box::new(gen(r, d - 1)), Box::new(gen(r, d - 1)), Box::new(gen(r, d - 1))); }
    let op = (rnd() % 4) as u8; E::B(op, Box::new(gen(r, d - 1)), Box::new(gen(r, d - 1)))
}
fn size(e: &E) -> usize { match e { E::V(_) | E::K(_) => 1, E::B(_, a, b) => 1 + size(a) + size(b), E::S(a, b, c, d) => 1 + size(a) + size(b) + size(c) + size(d) } }

fn llvm_expr<'c>(ctx: &'c Context, b: &inkwell::builder::Builder<'c>, e: &E, p: &[inkwell::values::FloatValue<'c>; 3]) -> inkwell::values::FloatValue<'c> {
    let f = ctx.f64_type();
    match e {
        E::V(i) => p[*i], E::K(k) => f.const_float(*k),
        E::B(op, a, c) => { let (x, y) = (llvm_expr(ctx, b, a, p), llvm_expr(ctx, b, c, p)); match op { 0 => b.build_float_add(x, y, "").unwrap(), 1 => b.build_float_sub(x, y, "").unwrap(), 2 => b.build_float_mul(x, y, "").unwrap(), _ => b.build_float_div(x, y, "").unwrap() } }
        E::S(a, c, t, u) => { let (x, y) = (llvm_expr(ctx, b, a, p), llvm_expr(ctx, b, c, p)); let cond = b.build_float_compare(FloatPredicate::OGT, x, y, "").unwrap(); let (tv, uv) = (llvm_expr(ctx, b, t, p), llvm_expr(ctx, b, u, p)); b.build_select(cond, tv, uv, "").unwrap().into_float_value() }
    }
}

fn cl_expr(b: &mut cranelift_frontend::FunctionBuilder, e: &E, p: &[cranelift_codegen::ir::Value; 3]) -> cranelift_codegen::ir::Value {
    use cranelift_codegen::ir::{condcodes::FloatCC, InstBuilder};
    match e {
        E::V(i) => p[*i], E::K(k) => b.ins().f64const(*k),
        E::B(op, a, c) => { let (x, y) = (cl_expr(b, a, p), cl_expr(b, c, p)); match op { 0 => b.ins().fadd(x, y), 1 => b.ins().fsub(x, y), 2 => b.ins().fmul(x, y), _ => b.ins().fdiv(x, y) } }
        E::S(a, c, t, u) => { let (x, y) = (cl_expr(b, a, p), cl_expr(b, c, p)); let cond = b.ins().fcmp(FloatCC::GreaterThan, x, y); let (tv, uv) = (cl_expr(b, t, p), cl_expr(b, u, p)); b.ins().select(cond, tv, uv) }
    }
}

fn scale_test() {
    let mut seed = 99u64;
    let max = 10_000; let exprs: Vec<E> = (0..max).map(|_| gen(&mut seed, 4)).collect();
    let avg = exprs.iter().map(size).sum::<usize>() as f64 / max as f64;
    println!("average formula size {avg:.1} nodes");
    for &n in &[100usize, 1000, 10000] {
        for (name, pipeline, lvl) in [("O0", "default<O0>", OptimizationLevel::None), ("O1", "default<O1>", OptimizationLevel::Less), ("O2", "default<O2>", OptimizationLevel::Default)] {
            let ctx = Context::create(); let t0 = Instant::now();
            let m = ctx.create_module("s"); let bld = ctx.create_builder(); let f64t = ctx.f64_type();
            let ft = f64t.fn_type(&[f64t.into(), f64t.into(), f64t.into()], false);
            for (k, e) in exprs[..n].iter().enumerate() {
                let f = m.add_function(&format!("f{k}"), ft, None); let bb = ctx.append_basic_block(f, "e"); bld.position_at_end(bb);
                let p = [f.get_nth_param(0).unwrap().into_float_value(), f.get_nth_param(1).unwrap().into_float_value(), f.get_nth_param(2).unwrap().into_float_value()];
                let v = llvm_expr(&ctx, &bld, e, &p); bld.build_return(Some(&v)).unwrap();
            }
            let t_ir = t0.elapsed(); let machine = tm(lvl);
            let t1 = Instant::now(); m.run_passes(pipeline, &machine, PassBuilderOptions::create()).unwrap(); let t_opt = t1.elapsed();
            let t2 = Instant::now(); let ee = m.create_jit_execution_engine(lvl).unwrap(); let addr = ee.get_function_address(&format!("f{}", n - 1)).unwrap(); let t_cg = t2.elapsed();
            std::hint::black_box(addr);
            let tot = t0.elapsed();
            println!("LLVM {name} N={n:6}: IR {:?} passes {:?} codegen {:?} total {:?} = {:.1} us/formula", t_ir, t_opt, t_cg, tot, tot.as_secs_f64() * 1e6 / n as f64);
        }
        use cranelift_codegen::ir::{types, AbiParam, InstBuilder};
        use cranelift_codegen::settings::{self, Configurable};
        use cranelift_frontend::{FunctionBuilder, FunctionBuilderContext};
        use cranelift_jit::{JITBuilder, JITModule};
        use cranelift_module::{Linkage, Module};
        for (name, opt) in [("none", "none"), ("speed", "speed")] {
            let t0 = Instant::now();
            let mut fl = settings::builder(); fl.set("opt_level", opt).unwrap();
            let isa = cranelift_native::builder().unwrap().finish(settings::Flags::new(fl)).unwrap();
            let mut m = JITModule::new(JITBuilder::with_isa(isa, cranelift_module::default_libcall_names()));
            let mut ctx = m.make_context(); let mut fbc = FunctionBuilderContext::new();
            let mut last = None;
            for (k, e) in exprs[..n].iter().enumerate() {
                ctx.clear();
                for _ in 0..3 { ctx.func.signature.params.push(AbiParam::new(types::F64)); }
                ctx.func.signature.returns.push(AbiParam::new(types::F64));
                let id = m.declare_function(&format!("f{k}"), Linkage::Export, &ctx.func.signature).unwrap();
                { let mut b = FunctionBuilder::new(&mut ctx.func, &mut fbc); let blk = b.create_block(); b.append_block_params_for_function_params(blk); b.switch_to_block(blk);
                  let ps = b.block_params(blk).to_vec(); let p = [ps[0], ps[1], ps[2]];
                  let v = cl_expr(&mut b, e, &p); b.ins().return_(&[v]); b.seal_all_blocks(); b.finalize(); }
                m.define_function(id, &mut ctx).unwrap(); last = Some(id);
            }
            m.finalize_definitions().unwrap(); let _ = m.get_finalized_function(last.unwrap());
            let tot = t0.elapsed();
            println!("Cranelift {name:5} N={n:6}: total {:?} = {:.1} us/formula", tot, tot.as_secs_f64() * 1e6 / n as f64);
        }
    }
}

fn cl_kernel() {
    use cranelift_codegen::ir::{condcodes::FloatCC, condcodes::IntCC, types, AbiParam, InstBuilder, MemFlags};
    use cranelift_codegen::settings::{self, Configurable};
    use cranelift_frontend::{FunctionBuilder, FunctionBuilderContext};
    use cranelift_jit::{JITBuilder, JITModule};
    use cranelift_module::{Linkage, Module};
    let n = 300_000usize; let (a, b, c) = data(n);
    for (label, rnd_exact) in [("approx round (nearest)", false), ("exact round via call", true)] {
        let t0 = Instant::now();
        let mut fl = settings::builder(); fl.set("opt_level", "speed").unwrap();
        let isa = cranelift_native::builder().unwrap().finish(settings::Flags::new(fl)).unwrap();
        let mut jb = JITBuilder::with_isa(isa, cranelift_module::default_libcall_names());
        jb.symbol("round2", round2 as *const u8);
        let mut m = JITModule::new(jb);
        let mut ctx = m.make_context(); let mut fbc = FunctionBuilderContext::new();
        let mut sig = m.make_signature(); sig.params.push(AbiParam::new(types::F64)); sig.returns.push(AbiParam::new(types::F64));
        let r2 = m.declare_function("round2", Linkage::Import, &sig).unwrap();
        ctx.func.signature.params.push(AbiParam::new(types::F64));
        for _ in 0..8 { ctx.func.signature.params.push(AbiParam::new(types::I64)); }
        let id = m.declare_function("kernel", Linkage::Export, &ctx.func.signature).unwrap();
        {
            let mut b_ = FunctionBuilder::new(&mut ctx.func, &mut fbc);
            let r2ref = m.declare_func_in_func(r2, b_.func);
            let entry = b_.create_block(); b_.append_block_params_for_function_params(entry);
            let lp = b_.create_block(); b_.append_block_param(lp, types::I64); b_.append_block_param(lp, types::F64);
            let done = b_.create_block();
            b_.switch_to_block(entry);
            let ps = b_.block_params(entry).to_vec();
            let (rate, nn, pa, pb, pc, pd, pe, pg, ph) = (ps[0], ps[1], ps[2], ps[3], ps[4], ps[5], ps[6], ps[7], ps[8]);
            let zi = b_.ins().iconst(types::I64, 0); let zf = b_.ins().f64const(0.0);
            b_.ins().jump(lp, &[zi.into(), zf.into()]);
            b_.switch_to_block(lp);
            let (i, run) = (b_.block_params(lp)[0], b_.block_params(lp)[1]);
            let off = b_.ins().ishl_imm(i, 3);
            let addr = |b_: &mut FunctionBuilder, base| b_.ins().iadd(base, off);
            let fl = MemFlags::trusted();
            let (xa, xb, xc) = { let q = addr(&mut b_, pa); let a_ = b_.ins().load(types::F64, fl, q, 0); let q = addr(&mut b_, pb); let b0 = b_.ins().load(types::F64, fl, q, 0); let q = addr(&mut b_, pc); let c_ = b_.ins().load(types::F64, fl, q, 0); (a_, b0, c_) };
            let x = b_.ins().fmul(xa, xb);
            let one = b_.ins().f64const(1.0); let om = b_.ins().fsub(one, rate); let disc = b_.ins().fmul(x, om);
            let th = b_.ins().f64const(50000.0); let gt = b_.ins().fcmp(FloatCC::GreaterThan, x, th);
            let a1 = b_.ins().select(gt, disc, x); let y = b_.ins().fmul(a1, xc);
            let d = if rnd_exact { let call = b_.ins().call(r2ref, &[y]); b_.inst_results(call)[0] } else { let h = b_.ins().f64const(100.0); let t = b_.ins().fmul(y, h); let r = b_.ins().nearest(t); b_.ins().fdiv(r, h) };
            let run2 = b_.ins().fadd(run, d); let q = b_.ins().fdiv(d, xa);
            let aq = b_.ins().fabs(q); let inf = b_.ins().f64const(f64::INFINITY); let fin = b_.ins().fcmp(FloatCC::LessThan, aq, inf);
            let z = b_.ins().f64const(0.0); let g = b_.ins().select(fin, q, z);
            let dg = b_.ins().fcmp(FloatCC::GreaterThan, d, g);
            let mx = b_.ins().select(dg, d, g); let mn = b_.ins().select(dg, g, d); let h = b_.ins().fsub(mx, mn);
            for (base, v) in [(pd, d), (pe, run2), (pg, g), (ph, h)] { let q = addr(&mut b_, base); b_.ins().store(fl, v, q, 0); }
            let inext = b_.ins().iadd_imm(i, 1);
            let cont = b_.ins().icmp(IntCC::UnsignedLessThan, inext, nn);
            b_.ins().brif(cont, lp, &[inext.into(), run2.into()], done, &[]);
            b_.switch_to_block(done); b_.ins().return_(&[]);
            b_.seal_all_blocks(); b_.finalize();
        }
        m.define_function(id, &mut ctx).unwrap(); m.finalize_definitions().unwrap();
        let f: Kern = unsafe { std::mem::transmute(m.get_finalized_function(id)) };
        let compile = t0.elapsed();
        let (mut d, mut e, mut g, mut h) = (vec![0.0; n], vec![0.0; n], vec![0.0; n], vec![0.0; n]);
        let mut best = f64::MAX;
        for _ in 0..7 { let t = Instant::now(); unsafe { f(0.05, n as i64, a.as_ptr(), b.as_ptr(), c.as_ptr(), d.as_mut_ptr(), e.as_mut_ptr(), g.as_mut_ptr(), h.as_mut_ptr()) }; best = best.min(t.elapsed().as_secs_f64()); }
        let want = reference(0.05, &a, &b, &c, rnd_exact);
        println!("Cranelift {label}: compile {:?}  run best {:.2} ms  identical_to_rust={}", compile, best * 1e3, rnd_exact && d == want.0 && e == want.1 && g == want.2 && h == want.3);
    }
}

fn wasm_obj() {
    Target::initialize_webassembly(&InitializationConfig::default());
    let triple = TargetTriple::create("wasm32-unknown-unknown"); let t = Target::from_triple(&triple).unwrap();
    let machine = t.create_target_machine(&triple, "generic", "", OptimizationLevel::Aggressive, RelocMode::PIC, CodeModel::Default).unwrap();
    for (exact, name, file) in [(true, "kernel_e", "kernel_e.o"), (false, "kernel_a", "kernel_a.o")] {
        let ctx = Context::create(); let m = llvm_kernel_named(&ctx, exact, name);
        m.set_triple(&triple); m.set_data_layout(&machine.get_target_data().get_data_layout());
        m.run_passes("default<O3>", &machine, PassBuilderOptions::create()).unwrap();
        machine.write_to_file(&m, FileType::Object, std::path::Path::new(file)).unwrap();
        println!("wrote {file}");
    }
}

fn bc_test() {
    unsafe { std::env::set_var("RT", "1") };
    let n = 300_000; let (a, b, c) = data(n);
    for link in [false, true] {
        let ctx = Context::create(); let m = llvm_kernel(&ctx, true);
        if link {
            let rt = inkwell::module::Module::parse_bitcode_from_path("../bc/rt.bc", &ctx).unwrap();
            m.link_in_module(rt).unwrap();
        }
        let machine = tm(OptimizationLevel::Aggressive);
        let t = Instant::now(); m.run_passes("default<O3>", &machine, PassBuilderOptions::create()).unwrap(); let tp = t.elapsed();
        let calls = m.print_to_string().to_string().matches("call double @rt_round2").count();
        let ee = m.create_jit_execution_engine(OptimizationLevel::Aggressive).unwrap();
        if let Some(r) = m.get_function("rt_round2") { if r.count_basic_blocks() == 0 { ee.add_global_mapping(&r, rt_round2 as usize); } }
        let f: Kern = unsafe { std::mem::transmute(ee.get_function_address("kernel").unwrap()) };
        let (mut d, mut e, mut g, mut h) = (vec![0.0; n], vec![0.0; n], vec![0.0; n], vec![0.0; n]);
        let mut best = f64::MAX;
        for _ in 0..7 { let t = Instant::now(); unsafe { f(0.05, n as i64, a.as_ptr(), b.as_ptr(), c.as_ptr(), d.as_mut_ptr(), e.as_mut_ptr(), g.as_mut_ptr(), h.as_mut_ptr()) }; best = best.min(t.elapsed().as_secs_f64()); }
        println!("LLVM 23, runtime leaf {}: passes {:?}, calls left in IR: {calls}, run best {:.2} ms, e[n-1]={}", if link { "LINKED from rustc bitcode (inlined)" } else { "called through ABI" }, tp, best * 1e3, e[n - 1]);
    }
}

fn main() {
    match std::env::args().nth(1).as_deref() { Some("kernel") => kernel_test(), Some("scale") => scale_test(), Some("cl") => cl_kernel(), Some("bc") => bc_test(), Some("wasmobj") => wasm_obj(), _ => eprintln!("kernel|scale|cl") }
}
