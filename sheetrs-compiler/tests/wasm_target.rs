//! The completion of a base module with a program (`sheetrs_compiler::wasm_target`), on a base made
//! here: one memory, one data segment holding the slot, and the rest as the real base has it. That a
//! real base completed answers requests is the test of `sheetrs-wasm` (Node).

use sheetrs_compiler::wasm_target::complete;
use wasm_encoder::{
    ConstExpr, DataSection, ExportKind, ExportSection, MemorySection, MemoryType, Module,
};

const SLOT: &[u8; 24] = b"SHEETRS-WASM-SLT\0\0\0\0\0\0\0\0";

/// A module with a memory of `pages` (and `max`), a segment of 100 bytes at 1 KiB that holds the slot
/// after 8 bytes, and a second segment after it.
fn base(pages: u64, max: Option<u64>) -> Vec<u8> {
    let mut module = Module::new();
    let mut memory = MemorySection::new();
    memory.memory(MemoryType {
        minimum: pages,
        maximum: max,
        memory64: false,
        shared: false,
        page_size_log2: None,
    });
    module.section(&memory);
    let mut exports = ExportSection::new();
    exports.export("memory", ExportKind::Memory, 0);
    module.section(&exports);
    let mut data = DataSection::new();
    let mut first = vec![7u8; 8];
    first.extend_from_slice(SLOT);
    first.extend(vec![9u8; 68]);
    data.active(0, &ConstExpr::i32_const(1024), first);
    data.active(0, &ConstExpr::i32_const(4096), vec![1u8, 2, 3]);
    module.section(&data);
    module.finish()
}

fn validate(bytes: &[u8]) {
    wasmparser::Validator::new()
        .validate_all(bytes)
        .expect("a valid module");
}

/// What a completed module has: its memory (minimum, maximum) and its data segments (offset, bytes).
/// A data segment: its offset and its bytes.
type Segment = (i32, Vec<u8>);

fn parse(bytes: &[u8]) -> ((u64, Option<u64>), Vec<Segment>) {
    use wasmparser::{DataKind, Operator, Parser, Payload};
    let mut memory = (0, None);
    let mut segments = Vec::new();
    for payload in Parser::new(0).parse_all(bytes) {
        match payload.expect("parses") {
            Payload::MemorySection(reader) => {
                for m in reader {
                    let m = m.unwrap();
                    memory = (m.initial, m.maximum);
                }
            }
            Payload::DataSection(reader) => {
                for d in reader {
                    let d = d.unwrap();
                    let DataKind::Active { offset_expr, .. } = d.kind else {
                        panic!("passive segment")
                    };
                    let Operator::I32Const { value } =
                        offset_expr.get_operators_reader().read().unwrap()
                    else {
                        panic!("not a constant offset")
                    };
                    segments.push((value, d.data.to_vec()));
                }
            }
            _ => {}
        }
    }
    (memory, segments)
}

#[test]
fn the_program_goes_after_the_memory_of_the_base_and_the_slot_points_to_it() {
    let program: Vec<u8> = (0..200_000u32).map(|i| (i % 251) as u8).collect();
    let done = complete(&base(2, None), &program).unwrap();
    validate(&done);
    let (memory, segments) = parse(&done);
    // The memory grew to hold the program, which starts where the memory of the base ended.
    assert_eq!(memory.0, (2 * 65_536 + 200_000u64).div_ceil(65_536));
    assert_eq!(segments.len(), 3);
    assert_eq!(segments[2].0, 2 * 65_536);
    assert_eq!(segments[2].1, program);
    // The slot says where, and the rest of the data of the base is as it was.
    let slot = &segments[0].1[8..32];
    assert_eq!(&slot[..16], &SLOT[..16]);
    assert_eq!(
        u32::from_le_bytes(slot[16..20].try_into().unwrap()),
        2 * 65_536
    );
    assert_eq!(
        u32::from_le_bytes(slot[20..24].try_into().unwrap()) as usize,
        program.len()
    );
    assert_eq!(&segments[0].1[..8], &[7u8; 8]);
    assert_eq!(&segments[0].1[32..], &[9u8; 68][..]);
    assert_eq!(segments[1], (4096, vec![1, 2, 3]));
}

#[test]
fn a_maximum_that_is_too_small_is_raised_with_the_memory() {
    let done = complete(&base(1, Some(1)), &vec![0u8; 100_000]).unwrap();
    validate(&done);
    let (memory, _) = parse(&done);
    assert_eq!(memory.0, 3);
    assert_eq!(memory.1, Some(3));
}

#[test]
fn a_module_that_has_a_program_is_not_a_base() {
    let done = complete(&base(1, None), b"program").unwrap();
    let error = complete(&done, b"another").unwrap_err().to_string();
    assert!(error.contains("already carries a program"), "{error}");
}

#[test]
fn what_is_not_a_base_is_said_so() {
    assert!(
        complete(b"not wasm", b"x")
            .unwrap_err()
            .to_string()
            .contains("not a WebAssembly")
    );
    // No slot: a module of the same shape without the marker.
    let mut module = Module::new();
    let mut memory = MemorySection::new();
    memory.memory(MemoryType {
        minimum: 1,
        maximum: None,
        memory64: false,
        shared: false,
        page_size_log2: None,
    });
    module.section(&memory);
    let mut data = DataSection::new();
    data.active(0, &ConstExpr::i32_const(0), vec![1u8; 10]);
    module.section(&data);
    let error = complete(&module.finish(), b"x").unwrap_err().to_string();
    assert!(error.contains("no slot"), "{error}");
}

#[test]
fn a_program_that_does_not_fit_is_refused() {
    let error = complete(&base(1, None), &vec![0u8; 3 << 30])
        .unwrap_err()
        .to_string();
    assert!(error.contains("too big"), "{error}");
}
