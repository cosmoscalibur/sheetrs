//! `OPT_BARRIER`, `OPT_TOUCH`, `OPT_PROB_HIT` and `OPT_PROB_INMONEY`, the pricing add-in of
//! LibreOffice (`scaddins/source/pricing`): barrier and touch options in the Black-Scholes model
//! with their greeks, as its `black_scholes.cxx` computes them. The results are checked against
//! the oracle (`tests/assets/semantics.*`).

use std::f64::consts::FRAC_1_SQRT_2;

use super::Reg;
use super::math_ext::erfc;
use super::num;
use crate::engine::{Engine, Res};
use crate::value::{ErrorKind, Value};

const DOMAIN: ErrorKind = ErrorKind::Domain;

#[derive(Clone, Copy, PartialEq)]
enum Greek {
    Value,
    Delta,
    Gamma,
    Theta,
    Vega,
    Volga,
    Vanna,
    RhoD,
    RhoF,
}

#[derive(Clone, Copy, PartialEq)]
enum PutCall {
    Call,
    Put,
}

#[derive(Clone, Copy, PartialEq)]
enum KoType {
    Regular,
    Reverse,
}

#[derive(Clone, Copy, PartialEq)]
enum Kio {
    KnockIn,
    KnockOut,
}

#[derive(Clone, Copy, PartialEq)]
enum Active {
    Continuous,
    Maturity,
}

#[derive(Clone, Copy, PartialEq)]
enum ForDom {
    Foreign,
    Domestic,
}

use Greek::{Delta, Gamma, RhoD, RhoF, Theta, Vanna, Vega, Volga};

fn sqr(x: f64) -> f64 {
    x * x
}

/// The normal density.
fn dnorm(x: f64) -> f64 {
    0.398_942_280_401_432_7 * (-0.5 * x * x).exp()
}

/// The normal distribution function.
fn pnorm(x: f64) -> f64 {
    0.5 * erfc(-x * FRAC_1_SQRT_2)
}

fn sign(pc: PutCall) -> f64 {
    if pc == PutCall::Call { 1.0 } else { -1.0 }
}

/// A binary option that pays 1 (domestic) in the case that `S_T` is above (call) or below (put)
/// the strike.
fn bincash(m: Model, k: f64, pc: PutCall, greek: Greek) -> f64 {
    let Model {
        s,
        vol,
        rd,
        rf,
        tau,
    } = m;
    let val = if tau <= 0.0 {
        match greek {
            Greek::Value if ((pc == PutCall::Call && s >= k) || (pc == PutCall::Put && s <= k)) => {
                1.0
            }
            _ => 0.0,
        }
    } else if k == 0.0 {
        if pc == PutCall::Put {
            0.0
        } else {
            match greek {
                Greek::Value => 1.0,
                Theta => rd,
                RhoD => -tau,
                _ => 0.0,
            }
        }
    } else {
        let d1 = ((s / k).ln() + (rd - rf + 0.5 * vol * vol) * tau) / (vol * tau.sqrt());
        let d2 = d1 - vol * tau.sqrt();
        let pm = sign(pc);
        match greek {
            Greek::Value => pnorm(pm * d2),
            Delta => pm * dnorm(d2) / (s * vol * tau.sqrt()),
            Gamma => -pm * dnorm(d2) * d1 / (sqr(s * vol) * tau),
            Theta => {
                rd * pnorm(pm * d2)
                    + pm * dnorm(d2) * ((s / k).ln() / (vol * tau.sqrt()) - 0.5 * d2) / tau
            }
            Vega => -pm * dnorm(d2) * d1 / vol,
            Volga => pm * dnorm(d2) / (vol * vol) * (-d1 * d1 * d2 + d1 + d2),
            Vanna => pm * dnorm(d2) / (s * vol * vol * tau.sqrt()) * (d1 * d2 - 1.0),
            RhoD => -tau * pnorm(pm * d2) + pm * dnorm(d2) * tau.sqrt() / vol,
            RhoF => -pm * dnorm(d2) * tau.sqrt() / vol,
        }
    };
    (-rd * tau).exp() * val
}

/// A binary option that pays `S_T` (foreign).
fn binasset(m: Model, k: f64, pc: PutCall, greek: Greek) -> f64 {
    let Model {
        s,
        vol,
        rd,
        rf,
        tau,
    } = m;
    let val = if tau <= 0.0 {
        let in_money = (pc == PutCall::Call && s >= k) || (pc == PutCall::Put && s <= k);
        match greek {
            Greek::Value => {
                if in_money {
                    s
                } else {
                    0.0
                }
            }
            Delta if in_money => 1.0,
            _ => 0.0,
        }
    } else if k == 0.0 {
        if pc == PutCall::Put {
            0.0
        } else {
            match greek {
                Greek::Value => s,
                Delta => 1.0,
                Theta => rf * s,
                RhoF => -tau * s,
                _ => 0.0,
            }
        }
    } else {
        let d1 = ((s / k).ln() + (rd - rf + 0.5 * vol * vol) * tau) / (vol * tau.sqrt());
        let d2 = d1 - vol * tau.sqrt();
        let pm = sign(pc);
        match greek {
            Greek::Value => s * pnorm(pm * d1),
            Delta => pnorm(pm * d1) + pm * dnorm(d1) / (vol * tau.sqrt()),
            Gamma => -pm * dnorm(d1) * d2 / (s * sqr(vol) * tau),
            Theta => {
                rf * s * pnorm(pm * d1)
                    + pm * s * dnorm(d1) * ((s / k).ln() / (vol * tau.sqrt()) - 0.5 * d1) / tau
            }
            Vega => -pm * s * dnorm(d1) * d2 / vol,
            Volga => pm * s * dnorm(d1) / (vol * vol) * (-d1 * d2 * d2 + d1 + d2),
            Vanna => pm * dnorm(d1) / (vol * vol * tau.sqrt()) * (d2 * d2 - 1.0),
            RhoD => pm * s * dnorm(d1) * tau.sqrt() / vol,
            RhoF => -tau * s * pnorm(pm * d1) - pm * s * dnorm(d1) * tau.sqrt() / vol,
        }
    };
    (-rf * tau).exp() * val
}

#[derive(Clone, Copy)]
struct Model {
    s: f64,
    vol: f64,
    rd: f64,
    rf: f64,
    tau: f64,
}

/// The terms of an option besides the model: the strike `k`, the lower barrier `b1` and the upper
/// `b2`, the payoff `pc` and the currency `fd` it is paid in.
#[derive(Clone, Copy)]
struct Contract {
    k: f64,
    b1: f64,
    b2: f64,
    pc: PutCall,
    fd: ForDom,
}

/// The binary option of the one currency or the other.
fn binary_k(m: Model, k: f64, pc: PutCall, fd: ForDom, greek: Greek) -> f64 {
    match fd {
        ForDom::Domestic => bincash(m, k, pc, greek),
        ForDom::Foreign => binasset(m, k, pc, greek),
    }
}

/// The binary option with a lower barrier `b1` and an upper `b2`, where one that is not above 0
/// is not set.
fn binary(m: Model, b1: f64, b2: f64, fd: ForDom, greek: Greek) -> f64 {
    if b1 <= 0.0 && b2 <= 0.0 {
        binary_k(m, 0.0, PutCall::Call, fd, greek)
    } else if b1 <= 0.0 && b2 > 0.0 {
        binary_k(m, b2, PutCall::Put, fd, greek)
    } else if b1 > 0.0 && b2 <= 0.0 {
        binary_k(m, b1, PutCall::Call, fd, greek)
    } else if b2 <= b1 {
        0.0
    } else {
        binary_k(m, b2, PutCall::Put, fd, greek) - binary_k(m, b1, PutCall::Put, fd, greek)
    }
}

/// The vanilla put or call.
fn putcall(m: Model, k: f64, pc: PutCall, greek: Greek) -> f64 {
    let (s, vol, rd, rf, tau) = (m.s, m.vol, m.rd, m.rf, m.tau);
    let pm = sign(pc);
    if k == 0.0 || tau == 0.0 {
        pm * (binasset(m, k, pc, greek) - k * bincash(m, k, pc, greek))
    } else {
        let d1 = ((s / k).ln() + (rd - rf + 0.5 * vol * vol) * tau) / (vol * tau.sqrt());
        let d2 = d1 - vol * tau.sqrt();
        match greek {
            Greek::Value => {
                pm * ((-rf * tau).exp() * s * pnorm(pm * d1)
                    - (-rd * tau).exp() * k * pnorm(pm * d2))
            }
            Delta => pm * (-rf * tau).exp() * pnorm(pm * d1),
            Gamma => (-rf * tau).exp() * dnorm(d1) / (s * vol * tau.sqrt()),
            _ => pm * (binasset(m, k, pc, greek) - k * bincash(m, k, pc, greek)),
        }
    }
}

/// The put or call truncated at one barrier.
fn putcalltrunc_one(m: Model, k: f64, b: f64, pc: PutCall, ko: KoType, greek: Greek) -> f64 {
    let pm = sign(pc);
    let degenerate = (pc == PutCall::Call && b <= k) || (pc == PutCall::Put && b >= k);
    match ko {
        KoType::Regular => {
            if degenerate {
                putcall(m, k, pc, greek)
            } else {
                pm * (binasset(m, b, pc, greek) - k * bincash(m, b, pc, greek))
            }
        }
        KoType::Reverse => {
            if degenerate {
                0.0
            } else {
                binasset(m, k, PutCall::Call, greek)
                    - binasset(m, b, PutCall::Call, greek)
                    - k * (bincash(m, k, PutCall::Call, greek)
                        - bincash(m, b, PutCall::Call, greek))
            }
        }
    }
}

/// The put or call truncated at none, one or two barriers.
fn putcalltrunc(m: Model, k: f64, b1: f64, b2: f64, pc: PutCall, greek: Greek) -> f64 {
    if b1 <= 0.0 && b2 <= 0.0 {
        putcall(m, k, pc, greek)
    } else if b1 <= 0.0 && b2 > 0.0 {
        let ko = if pc == PutCall::Call {
            KoType::Reverse
        } else {
            KoType::Regular
        };
        putcalltrunc_one(m, k, b2, pc, ko, greek)
    } else if b1 > 0.0 && b2 <= 0.0 {
        let ko = if pc == PutCall::Call {
            KoType::Regular
        } else {
            KoType::Reverse
        };
        putcalltrunc_one(m, k, b1, pc, ko, greek)
    } else if b2 <= b1 {
        0.0
    } else {
        sign(pc)
            * (putcalltrunc_one(m, k, b1, pc, KoType::Regular, greek)
                - putcalltrunc_one(m, k, b2, pc, KoType::Regular, greek))
    }
}

/// The option that does not depend on the path; a negative strike stands for a binary option.
fn vanilla(m: Model, k: f64, b1: f64, b2: f64, pc: PutCall, fd: ForDom, greek: Greek) -> f64 {
    if k < 0.0 {
        binary(m, b1, b2, fd, greek)
    } else {
        putcall(m, k, pc, greek)
    }
}

fn vanilla_trunc(m: Model, k: f64, b1: f64, b2: f64, pc: PutCall, fd: ForDom, greek: Greek) -> f64 {
    if k < 0.0 {
        binary(m, b1, b2, fd, greek)
    } else {
        putcalltrunc(m, k, b1, b2, pc, greek)
    }
}

/// What a barrier option is made of: `V(sc S) - (B/S)^a V(sc B^2/S)`, with `V` the vanilla
/// truncated at the barriers.
fn barrier_term(m: Model, contract: Contract, sc: f64, greek: Greek) -> f64 {
    let Contract { k, b1, b2, pc, fd } = contract;
    let (s, vol, rd, rf) = (m.s, m.vol, m.rd, m.rf);
    let b = if b1 > 0.0 { b1 } else { b2 };
    let a = 2.0 * (rd - rf) / (vol * vol) - 1.0;
    let bb = 4.0 * (rd - rf) / (vol * vol * vol);
    let c = 12.0 * (rd - rf) / (vol * vol * vol * vol);
    let at = |s2: f64, g: Greek| vanilla_trunc(Model { s: s2, ..m }, k, b1, b2, pc, fd, g);
    let pw = (b / s).powf(a);
    match greek {
        Greek::Value | Theta => at(sc * s, greek) - pw * at(sc * b * b / s, greek),
        Delta => {
            sc * at(sc * s, greek)
                + pw * (a / s * at(sc * b * b / s, Greek::Value)
                    + sqr(b / s) * sc * at(sc * b * b / s, greek))
        }
        Gamma => {
            sc * sc * at(sc * s, greek)
                - pw * (a * (a + 1.0) / (s * s) * at(sc * b * b / s, Greek::Value)
                    + (2.0 * a + 2.0) * b * b / (s * s * s) * sc * at(sc * b * b / s, Delta)
                    + sqr(sqr(b / s)) * sc * sc * at(sc * b * b / s, Gamma))
        }
        Vega => {
            at(sc * s, greek)
                - pw * (-bb * (b / s).ln() * at(sc * b * b / s, Greek::Value)
                    + 1.0 * at(sc * b * b / s, greek))
        }
        Volga => {
            at(sc * s, greek)
                - pw * ((b / s).ln()
                    * (bb * bb * (b / s).ln() + c)
                    * at(sc * b * b / s, Greek::Value)
                    - 2.0 * bb * (b / s).ln() * at(sc * b * b / s, Vega)
                    + 1.0 * at(sc * b * b / s, Volga))
        }
        Vanna => {
            sc * at(sc * s, greek)
                - pw * (bb / s * ((b / s).ln() * a + 1.0) * at(b * b / s * sc, Greek::Value)
                    + bb * (b / s).ln() * sqr(b / s) * sc * at(b * b / s * sc, Delta)
                    - a / s * at(b * b / s * sc, Vega)
                    - sqr(b / s) * sc * at(b * b / s * sc, Vanna))
        }
        RhoD => {
            at(sc * s, greek)
                - pw * (2.0 * (b / s).ln() / (vol * vol) * at(sc * b * b / s, Greek::Value)
                    + 1.0 * at(sc * b * b / s, greek))
        }
        RhoF => {
            at(sc * s, greek)
                - pw * (-2.0 * (b / s).ln() / (vol * vol) * at(sc * b * b / s, Greek::Value)
                    + 1.0 * at(sc * b * b / s, greek))
        }
    }
}

/// A term of the infinite sum of a double barrier.
fn barrier_double_term(
    m: Model,
    contract: Contract,
    fac: f64,
    sc: f64,
    i: i32,
    greek: Greek,
) -> f64 {
    let Contract { b1, b2, .. } = contract;
    let (vol, rd, rf) = (m.vol, m.rd, m.rf);
    let i = f64::from(i);
    let b = 4.0 * i * (rd - rf) / (vol * vol * vol);
    let c = 12.0 * i * (rd - rf) / (vol * vol * vol * vol);
    let term = |g: Greek| barrier_term(m, contract, sc, g);
    match greek {
        Greek::Value | Delta | Gamma | Theta => fac * term(greek),
        Vega => fac * term(greek) - b * (b2 / b1).ln() * fac * term(Greek::Value),
        Volga => {
            fac * term(greek) - 2.0 * b * (b2 / b1).ln() * fac * term(Vega)
                + (b2 / b1).ln() * fac * (c + b * b * (b2 / b1).ln()) * term(Greek::Value)
        }
        Vanna => fac * term(greek) - b * (b2 / b1).ln() * fac * term(Delta),
        RhoD => {
            fac * term(greek) + 2.0 * i / (vol * vol) * (b2 / b1).ln() * fac * term(Greek::Value)
        }
        RhoF => {
            fac * term(greek) - 2.0 * i / (vol * vol) * (b2 / b1).ln() * fac * term(Greek::Value)
        }
    }
}

/// The knock-out option with no, one or two barriers monitored continuously.
fn barrier_ko(m: Model, k: f64, b1: f64, b2: f64, pc: PutCall, fd: ForDom, greek: Greek) -> f64 {
    let contract = Contract { k, b1, b2, pc, fd };
    if b1 <= 0.0 && b2 <= 0.0 {
        vanilla(m, k, b1, b2, pc, fd, greek)
    } else if b1 > 0.0 && b2 <= 0.0 {
        if m.s <= b1 {
            0.0
        } else {
            barrier_term(m, contract, 1.0, greek)
        }
    } else if b1 <= 0.0 && b2 > 0.0 {
        if m.s >= b2 {
            0.0
        } else {
            barrier_term(m, contract, 1.0, greek)
        }
    } else if m.s <= b1 || m.s >= b2 {
        0.0
    } else {
        let a = 2.0 * (m.rd - m.rf) / (m.vol * m.vol) - 1.0;
        let bb2 = sqr(b2 / b1);
        let bba = (b2 / b1).powf(a);
        let bb2inv = 1.0 / bb2;
        let bbainv = 1.0 / bba;
        let (mut fac, mut facinv, mut sc, mut scinv) = (1.0, 1.0, 1.0, 1.0);
        let mut val = barrier_double_term(m, contract, fac, sc, 0, greek);
        for i in 1..10 {
            fac *= bba;
            facinv *= bbainv;
            sc *= bb2;
            scinv *= bb2inv;
            let add = barrier_double_term(m, contract, fac, sc, i, greek)
                + barrier_double_term(m, contract, facinv, scinv, -i, greek);
            val += add;
            if add.abs() <= 1e-12 * val.abs() {
                break;
            }
        }
        val
    }
}

fn barrier_ki(m: Model, k: f64, b1: f64, b2: f64, pc: PutCall, fd: ForDom, greek: Greek) -> f64 {
    vanilla(m, k, b1, b2, pc, fd, greek) - barrier_ko(m, k, b1, b2, pc, fd, greek)
}

fn barrier_general(m: Model, contract: Contract, kio: Kio, active: Active, greek: Greek) -> f64 {
    let Contract { k, b1, b2, pc, fd } = contract;
    match (kio, active) {
        (Kio::KnockOut, Active::Maturity) => vanilla_trunc(m, k, b1, b2, pc, fd, greek),
        (Kio::KnockOut, Active::Continuous) => barrier_ko(m, k, b1, b2, pc, fd, greek),
        (Kio::KnockIn, Active::Maturity) => {
            vanilla(m, k, b1, b2, pc, fd, greek) - vanilla_trunc(m, k, b1, b2, pc, fd, greek)
        }
        (Kio::KnockIn, Active::Continuous) => barrier_ki(m, k, b1, b2, pc, fd, greek),
    }
}

/// A touch or no-touch option.
fn touch(m: Model, b1: f64, b2: f64, fd: ForDom, kio: Kio, active: Active, greek: Greek) -> f64 {
    let k = -1.0;
    let pc = PutCall::Call;
    match (kio, active) {
        (Kio::KnockOut, Active::Maturity) => vanilla_trunc(m, k, b1, b2, pc, fd, greek),
        (Kio::KnockOut, Active::Continuous) => barrier_ko(m, k, b1, b2, pc, fd, greek),
        (Kio::KnockIn, Active::Maturity) => {
            vanilla(m, k, -1.0, -1.0, pc, fd, greek) - vanilla_trunc(m, k, b1, b2, pc, fd, greek)
        }
        (Kio::KnockIn, Active::Continuous) => {
            vanilla(m, k, -1.0, -1.0, pc, fd, greek) - barrier_ko(m, k, b1, b2, pc, fd, greek)
        }
    }
}

/// A barrier option with a put or call payoff, and its rebate.
fn barrier(
    m: Model,
    contract: Contract,
    rebate: f64,
    kio: Kio,
    active: Active,
    greek: Greek,
) -> f64 {
    let Contract { b1, b2, fd, .. } = contract;
    let mut val = barrier_general(m, contract, kio, active, greek);
    if rebate != 0.0 {
        let other = if kio == Kio::KnockIn {
            Kio::KnockOut
        } else {
            Kio::KnockIn
        };
        val += rebate * touch(m, b1, b2, fd, other, active, greek);
    }
    val
}

fn prob_hit(s: f64, vol: f64, mu: f64, tau: f64, b1: f64, b2: f64) -> f64 {
    let m = Model {
        s,
        vol,
        rd: 0.0,
        rf: -mu,
        tau,
    };
    1.0 - touch(
        m,
        b1,
        b2,
        ForDom::Domestic,
        Kio::KnockOut,
        Active::Continuous,
        Greek::Value,
    )
}

fn prob_in_money_barriers(m: Model, b1: f64, b2: f64) -> f64 {
    if b1 < b2 || b1 <= 0.0 || b2 <= 0.0 {
        binary(m, b1, b2, ForDom::Domestic, Greek::Value)
    } else {
        0.0
    }
}

fn prob_in_money(m: Model, k: f64, b1: f64, b2: f64, pc: PutCall) -> f64 {
    if k < 0.0 {
        return prob_in_money_barriers(m, b1, b2);
    }
    if (b1 > b2 && b1 > 0.0 && b2 > 0.0)
        || (k >= b2 && b2 > 0.0 && pc == PutCall::Call)
        || (k <= b1 && pc == PutCall::Put)
    {
        0.0
    } else if pc == PutCall::Call {
        prob_in_money_barriers(m, b1.max(k), b2)
    } else {
        let bm2 = if b2 > 0.0 { b2.min(k) } else { k };
        prob_in_money_barriers(m, b1, bm2)
    }
}

// The arguments, as the add-in reads them.

/// An argument that is a text: a number is its text.
fn string_arg(e: &mut Engine, v: &Value) -> Res<String> {
    e.text(v)
}

fn put_call(s: &str) -> Res<PutCall> {
    if s.starts_with('c') {
        Ok(PutCall::Call)
    } else if s.starts_with('p') {
        Ok(PutCall::Put)
    } else {
        Err(DOMAIN)
    }
}

fn in_out(s: &str) -> Res<Kio> {
    if s.starts_with('i') {
        Ok(Kio::KnockIn)
    } else if s.starts_with('o') {
        Ok(Kio::KnockOut)
    } else {
        Err(DOMAIN)
    }
}

fn active(s: &str) -> Res<Active> {
    if s.starts_with('c') {
        Ok(Active::Continuous)
    } else if s.starts_with('e') {
        Ok(Active::Maturity)
    } else {
        Err(DOMAIN)
    }
}

fn for_dom(s: &str) -> Res<ForDom> {
    if s.starts_with('f') {
        Ok(ForDom::Foreign)
    } else if s.starts_with('d') {
        Ok(ForDom::Domestic)
    } else {
        Err(DOMAIN)
    }
}

/// An optional argument that may be a text, or missing.
fn any_text(e: &mut Engine, a: &[Value], i: usize) -> Res<Option<String>> {
    match a.get(i).map(|v| e.scalar(v.clone())) {
        None | Some(Value::Empty) => Ok(None),
        Some(Value::Error(k)) => Err(k),
        Some(Value::Text(t)) => Ok(Some(t.to_string())),
        Some(_) => Err(DOMAIN),
    }
}

fn greek(e: &mut Engine, a: &[Value], i: usize) -> Res<Greek> {
    let s = any_text(e, a, i)?.unwrap_or_else(|| "value".to_string());
    Ok(match s.as_str() {
        "value" | "price" | "v" | "p" => Greek::Value,
        "delta" | "d" => Delta,
        "gamma" | "g" => Gamma,
        "theta" | "t" => Theta,
        "vega" | "e" => Vega,
        "volga" | "o" => Volga,
        "vanna" | "a" => Vanna,
        "rho" | "r" => RhoD,
        "rhof" | "f" => RhoF,
        _ => return Err(DOMAIN),
    })
}

fn finite(x: f64) -> Res<Value> {
    if x.is_finite() { num(x) } else { Err(DOMAIN) }
}

/// The first of the arguments that is an error is the result.
fn first_error(e: &mut Engine, a: &[Value]) -> Res<()> {
    for v in a {
        if let Value::Error(k) = e.scalar(v.clone()) {
            return Err(k);
        }
    }
    Ok(())
}

pub(super) fn register(r: &mut Reg) {
    r.eager("OPT_BARRIER", 12, 13, |e, a| {
        first_error(e, a)?;
        let n: Vec<f64> = (0..9).map(|i| e.num(&a[i])).collect::<Res<_>>()?;
        let (spot, vol, rd, rf, tau, strike, b1, b2, rebate) =
            (n[0], n[1], n[2], n[3], n[4], n[5], n[6], n[7], n[8]);
        let pc = string_arg(e, &a[9])?;
        let kio = string_arg(e, &a[10])?;
        let act = string_arg(e, &a[11])?;
        if spot <= 0.0 || vol <= 0.0 || tau < 0.0 || strike < 0.0 {
            return Err(DOMAIN);
        }
        let (pc, kio, act) = (put_call(&pc)?, in_out(&kio)?, active(&act)?);
        let g = greek(e, a, 12)?;
        finite(barrier(
            Model {
                s: spot,
                vol,
                rd,
                rf,
                tau,
            },
            Contract {
                k: strike,
                b1,
                b2,
                pc,
                fd: ForDom::Domestic,
            },
            rebate,
            kio,
            act,
            g,
        ))
    });
    r.eager("OPT_TOUCH", 10, 11, |e, a| {
        first_error(e, a)?;
        let n: Vec<f64> = (0..7).map(|i| e.num(&a[i])).collect::<Res<_>>()?;
        let (spot, vol, rd, rf, tau, b1, b2) = (n[0], n[1], n[2], n[3], n[4], n[5], n[6]);
        let fd = string_arg(e, &a[7])?;
        let kio = string_arg(e, &a[8])?;
        let act = string_arg(e, &a[9])?;
        if spot <= 0.0 || vol <= 0.0 || tau < 0.0 {
            return Err(DOMAIN);
        }
        let (fd, kio, act) = (for_dom(&fd)?, in_out(&kio)?, active(&act)?);
        let g = greek(e, a, 10)?;
        finite(touch(
            Model {
                s: spot,
                vol,
                rd,
                rf,
                tau,
            },
            b1,
            b2,
            fd,
            kio,
            act,
            g,
        ))
    });
    r.eager("OPT_PROB_HIT", 6, 6, |e, a| {
        first_error(e, a)?;
        let n: Vec<f64> = (0..6).map(|i| e.num(&a[i])).collect::<Res<_>>()?;
        if n[0] <= 0.0 || n[1] <= 0.0 || n[3] < 0.0 {
            return Err(DOMAIN);
        }
        finite(prob_hit(n[0], n[1], n[2], n[3], n[4], n[5]))
    });
    r.eager("OPT_PROB_INMONEY", 6, 8, |e, a| {
        first_error(e, a)?;
        let n: Vec<f64> = (0..6).map(|i| e.num(&a[i])).collect::<Res<_>>()?;
        let (spot, vol, mu, tau, b1, b2) = (n[0], n[1], n[2], n[3], n[4], n[5]);
        if spot <= 0.0 || vol <= 0.0 || tau < 0.0 {
            return Err(DOMAIN);
        }
        // The put or call is read first (a text starting with `c` or `p`, `c` when missing), then
        // the strike (a number, -1 when missing).
        let pc = match any_text(e, a, 7)? {
            None => PutCall::Call,
            Some(t) => put_call(&t)?,
        };
        let k = match a.get(6).map(|v| e.scalar(v.clone())) {
            None | Some(Value::Empty) => -1.0,
            Some(Value::Number(x)) => x,
            Some(Value::Bool(b)) => f64::from(u8::from(b)),
            Some(_) => return Err(DOMAIN),
        };
        let m = Model {
            s: spot,
            vol,
            rd: 0.0,
            rf: -mu,
            tau,
        };
        finite(prob_in_money(m, k, b1, b2, pc))
    });
}
