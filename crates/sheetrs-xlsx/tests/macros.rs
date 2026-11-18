//! VBA projects, built from the public layouts of the compound file and of the project, and
//! the macros of a workbook that holds one.

use std::error::Error as StdError;
use std::fmt::Write as _;
use std::fs::File;
use std::io::Cursor;

use sheetrs_core::{Cell, MacroKind};
use sheetrs_testkit::case::CellValue;
use sheetrs_testkit::compare::Leniency;
use sheetrs_testkit::crafted::{self, Package};

mod common;
use common::{REL_NS, rels_xml, single_sheet_package_with, workbook_xml, worksheet_xml};

const DIR: &str = env!("CARGO_TARGET_TMPDIR");
const END_OF_CHAIN: u32 = 0xFFFF_FFFE;
const FREE: u32 = 0xFFFF_FFFF;
const FAT_SECTOR: u32 = 0xFFFF_FFFD;

/// A compressed container of `text`. A chunk of at most 3640 bytes is kept as literals, each
/// group of eight bytes after a flag byte of zeros; a longer one does not fit the size of a
/// chunk as literals, so it is kept as it is, padded to 4096 bytes.
fn compress(text: &[u8]) -> Vec<u8> {
    let mut out = vec![1];
    for chunk in text.chunks(4096) {
        if chunk.len() <= 3640 {
            let mut data = Vec::new();
            for group in chunk.chunks(8) {
                data.push(0);
                data.extend_from_slice(group);
            }
            let size = u16::try_from(data.len() + 2 - 3).unwrap_or(0);
            out.extend_from_slice(&(0xB000 | size).to_le_bytes());
            out.extend_from_slice(&data);
        } else {
            out.extend_from_slice(&0x3FFF_u16.to_le_bytes());
            let mut padded = chunk.to_vec();
            padded.resize(4096, 0);
            out.extend_from_slice(&padded);
        }
    }
    out
}

/// A record of the `dir` stream: its id, the size of its data and the data.
fn record(id: u16, data: &[u8]) -> Vec<u8> {
    let mut out = id.to_le_bytes().to_vec();
    out.extend_from_slice(&u32::try_from(data.len()).unwrap_or(0).to_le_bytes());
    out.extend_from_slice(data);
    out
}

fn utf16(text: &str) -> Vec<u8> {
    text.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

/// The reference records of a project that refers to a standard library and to a control, in the
/// layouts the format gives them.
fn reference_records() -> Vec<u8> {
    let libid = b"*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\stdole2.tlb#OLE Automation";
    let length = |bytes: &[u8]| u32::try_from(bytes.len()).unwrap_or(0).to_le_bytes();
    let counted = |bytes: &[u8]| [length(bytes).as_slice(), bytes].concat();
    let mut out = Vec::new();
    out.extend(record(0x0016, b"stdole"));
    out.extend(record(0x003E, &utf16("stdole")));
    out.extend(record(
        0x000D,
        &[
            counted(libid),
            0_u32.to_le_bytes().to_vec(),
            0_u16.to_le_bytes().to_vec(),
        ]
        .concat(),
    ));
    out.extend(record(0x0033, &counted(libid)));
    out.extend(record(
        0x000E,
        &[
            counted(libid),
            counted(b"*\\RProject"),
            1_u32.to_le_bytes().to_vec(),
            0_u16.to_le_bytes().to_vec(),
        ]
        .concat(),
    ));
    // A control reference goes on after the size it gives: its name, then its extended part.
    out.extend(record(
        0x002F,
        &[
            counted(libid),
            0_u32.to_le_bytes().to_vec(),
            0_u16.to_le_bytes().to_vec(),
        ]
        .concat(),
    ));
    out.extend(record(0x0016, b"control"));
    out.extend(record(0x003E, &utf16("control")));
    out.extend(record(
        0x0030,
        &[
            counted(libid),
            0_u32.to_le_bytes().to_vec(),
            0_u16.to_le_bytes().to_vec(),
            vec![0_u8; 16],
            0_u32.to_le_bytes().to_vec(),
        ]
        .concat(),
    ));
    out
}

/// The `dir` stream of a project called `name` with `modules`, which is not compressed yet.
fn dir_stream(name: &str, modules: &[&str], references: bool) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend(record(0x0001, &1_u32.to_le_bytes()));
    out.extend(record(0x0002, &0x0409_u32.to_le_bytes()));
    out.extend(record(0x0014, &0x0409_u32.to_le_bytes()));
    out.extend(record(0x0003, &1252_u16.to_le_bytes()));
    out.extend(record(0x0004, name.as_bytes()));
    out.extend(record(0x0005, b""));
    out.extend(record(0x0040, b""));
    out.extend(record(0x0006, b""));
    out.extend(record(0x003D, b""));
    out.extend(record(0x0007, &0_u32.to_le_bytes()));
    out.extend(record(0x0008, &0_u32.to_le_bytes()));
    // The version record says its size is 4 and goes on with the major version and 2 more
    // bytes for the minor one.
    out.extend(0x0009_u16.to_le_bytes());
    out.extend(4_u32.to_le_bytes());
    out.extend(1_u32.to_le_bytes());
    out.extend(0_u16.to_le_bytes());
    out.extend(record(0x000C, b""));
    out.extend(record(0x003C, b""));
    if references {
        out.extend(reference_records());
    }
    out.extend(record(
        0x000F,
        &u16::try_from(modules.len()).unwrap_or(0).to_le_bytes(),
    ));
    out.extend(record(0x0013, &0xFFFF_u16.to_le_bytes()));
    for module in modules {
        out.extend(record(0x0019, module.as_bytes()));
        out.extend(record(0x0047, &utf16(module)));
        out.extend(record(0x001A, module.as_bytes()));
        out.extend(record(0x0032, &utf16(module)));
        out.extend(record(0x001C, b""));
        out.extend(record(0x0048, b""));
        out.extend(record(0x0031, &0_u32.to_le_bytes()));
        out.extend(record(0x001E, &0_u32.to_le_bytes()));
        out.extend(record(0x002C, &0xFFFF_u16.to_le_bytes()));
        out.extend(record(0x0021, b""));
        out.extend(record(0x002B, b""));
    }
    out.extend(record(0x0010, b""));
    out
}

/// The `PROJECT` stream of a project called `name` that lists `modules`.
fn project_stream(name: &str, modules: &[&str]) -> Vec<u8> {
    let mut text = String::from("ID=\"{00000000-0000-0000-0000-000000000000}\"\r\n");
    for module in modules {
        let _ = write!(text, "Module={module}\r\n");
    }
    let _ = write!(text, "Name=\"{name}\"\r\nHelpContextID=\"0\"\r\n");
    text.into_bytes()
}

/// A stream or storage to place in a compound file.
enum Node {
    Storage(&'static str, Vec<Node>),
    Stream(String, Vec<u8>),
}

fn node_name(node: &Node) -> String {
    match node {
        Node::Storage(name, _) => (*name).to_owned(),
        Node::Stream(name, _) => name.clone(),
    }
}

/// An entry of the directory of a compound file being built.
struct Slot {
    name: String,
    kind: u8,
    right: u32,
    child: u32,
    start: u32,
    size: u32,
}

/// The directory and the mini stream of a compound file being built.
struct Builder {
    slots: Vec<Slot>,
    mini: Vec<u8>,
    mini_fat: Vec<u32>,
    /// The streams too big for the mini stream, with the entry each belongs to.
    large: Vec<(usize, Vec<u8>)>,
}

impl Builder {
    fn index(&self) -> u32 {
        u32::try_from(self.slots.len()).unwrap_or(FREE)
    }

    /// Places a stream in the mini stream and returns the mini sector it starts at.
    fn store(&mut self, data: &[u8]) -> u32 {
        let start = u32::try_from(self.mini.len() / 64).unwrap_or(0);
        let sectors = data.len().div_ceil(64).max(1);
        for number in 0..sectors {
            self.mini_fat.push(if number + 1 < sectors {
                start + u32::try_from(number + 1).unwrap_or(0)
            } else {
                END_OF_CHAIN
            });
        }
        let mut padded = data.to_vec();
        padded.resize(sectors * 64, 0);
        self.mini.extend_from_slice(&padded);
        start
    }

    /// Places the siblings in a chain, ordered the way compound files order names, and
    /// returns the index of the first.
    fn place(&mut self, mut siblings: Vec<Node>) -> u32 {
        siblings.sort_by_key(|node| {
            let name = node_name(node);
            (name.len(), name.to_uppercase())
        });
        let first = self.index();
        let count = siblings.len();
        let mut storages = Vec::new();
        for (position, node) in siblings.into_iter().enumerate() {
            let index = self.index();
            let right = if position + 1 < count {
                index + 1
            } else {
                FREE
            };
            match node {
                Node::Storage(name, inner) => {
                    self.slots.push(Slot {
                        name: name.to_owned(),
                        kind: 1,
                        right,
                        child: FREE,
                        start: 0,
                        size: 0,
                    });
                    storages.push((index, inner));
                }
                Node::Stream(name, data) if data.len() >= 4096 => {
                    self.large
                        .push((usize::try_from(index).unwrap_or(0), data.clone()));
                    self.slots.push(Slot {
                        name,
                        kind: 2,
                        right,
                        child: FREE,
                        start: 0,
                        size: u32::try_from(data.len()).unwrap_or(0),
                    });
                }
                Node::Stream(name, data) => {
                    let start = self.store(&data);
                    self.slots.push(Slot {
                        name,
                        kind: 2,
                        right,
                        child: FREE,
                        start,
                        size: u32::try_from(data.len()).unwrap_or(0),
                    });
                }
            }
        }
        // The children of a storage are placed after all its siblings.
        for (parent, inner) in storages {
            let child = self.place(inner);
            if let Some(slot) = self.slots.get_mut(usize::try_from(parent).unwrap_or(0)) {
                slot.child = child;
            }
        }
        first
    }

    /// The directory as sectors of four entries.
    fn directory(&self) -> Vec<u8> {
        let mut out = Vec::new();
        for number in 0..self.slots.len().div_ceil(4) * 4 {
            let mut bytes = vec![0_u8; 128];
            for field in [68, 72, 76] {
                bytes[field..field + 4].copy_from_slice(&FREE.to_le_bytes());
            }
            if let Some(slot) = self.slots.get(number) {
                let units = utf16(&slot.name);
                bytes[..units.len()].copy_from_slice(&units);
                let length = u16::try_from(units.len() + 2).unwrap_or(0);
                bytes[64..66].copy_from_slice(&length.to_le_bytes());
                bytes[66] = slot.kind;
                bytes[67] = 1;
                bytes[72..76].copy_from_slice(&slot.right.to_le_bytes());
                bytes[76..80].copy_from_slice(&slot.child.to_le_bytes());
                bytes[116..120].copy_from_slice(&slot.start.to_le_bytes());
                bytes[120..124].copy_from_slice(&slot.size.to_le_bytes());
            }
            out.extend_from_slice(&bytes);
        }
        out
    }
}

/// The header of a compound file whose allocation table takes the first `fat_sectors` sectors,
/// followed by the directory, the mini table, the mini stream and the large streams.
fn header(fat_sectors: usize, directory_sectors: usize, mini_fat_sectors: usize) -> Vec<u8> {
    let first_directory = u32::try_from(fat_sectors).unwrap_or(0);
    let first_mini_fat = first_directory + u32::try_from(directory_sectors).unwrap_or(0);
    let mut out = vec![0_u8; 512];
    out[..8].copy_from_slice(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]);
    out[24..26].copy_from_slice(&0x003E_u16.to_le_bytes());
    out[26..28].copy_from_slice(&3_u16.to_le_bytes());
    out[28..30].copy_from_slice(&0xFFFE_u16.to_le_bytes());
    out[30..32].copy_from_slice(&9_u16.to_le_bytes());
    out[32..34].copy_from_slice(&6_u16.to_le_bytes());
    out[44..48].copy_from_slice(&first_directory.to_le_bytes());
    out[48..52].copy_from_slice(&first_directory.to_le_bytes());
    out[56..60].copy_from_slice(&4096_u32.to_le_bytes());
    out[60..64].copy_from_slice(&first_mini_fat.to_le_bytes());
    out[64..68].copy_from_slice(&u32::try_from(mini_fat_sectors).unwrap_or(0).to_le_bytes());
    out[68..72].copy_from_slice(&END_OF_CHAIN.to_le_bytes());
    for index in 0..109 {
        let value = if index < fat_sectors {
            u32::try_from(index).unwrap_or(FREE)
        } else {
            FREE
        };
        out[76 + 4 * index..80 + 4 * index].copy_from_slice(&value.to_le_bytes());
    }
    out
}

/// The allocation table of `fat_sectors` sectors, for a file whose sectors after it are `chains`,
/// each a run of that many sectors.
fn allocation_table(fat_sectors: usize, chains: &[usize]) -> Vec<u8> {
    let mut fat = vec![FREE; fat_sectors * 128];
    fat[..fat_sectors].fill(FAT_SECTOR);
    let mut at = fat_sectors;
    for count in chains {
        for number in 0..*count {
            fat[at + number] = if number + 1 < *count {
                u32::try_from(at + number + 1).unwrap_or(END_OF_CHAIN)
            } else {
                END_OF_CHAIN
            };
        }
        at += count;
    }
    fat.iter().flat_map(|entry| entry.to_le_bytes()).collect()
}

/// A compound file of 512-byte sectors with `children` in its root.
fn compound_file(children: Vec<Node>) -> Vec<u8> {
    let mut builder = Builder {
        slots: vec![Slot {
            name: "Root Entry".to_owned(),
            kind: 5,
            right: FREE,
            child: FREE,
            start: END_OF_CHAIN,
            size: 0,
        }],
        mini: Vec::new(),
        mini_fat: Vec::new(),
        large: Vec::new(),
    };
    let root_child = builder.place(children);
    let directory_sectors = builder.slots.len().div_ceil(4);
    let mini_fat_sectors = builder.mini_fat.len().div_ceil(128).max(1);
    let mini_sectors = builder.mini.len().div_ceil(512).max(1);
    let large_sectors: Vec<usize> = builder
        .large
        .iter()
        .map(|(_, data)| data.len().div_ceil(512))
        .collect();
    let others =
        directory_sectors + mini_fat_sectors + mini_sectors + large_sectors.iter().sum::<usize>();
    // The table lists itself, so each of its sectors serves 127 others.
    let fat_sectors = others.div_ceil(127);
    let mut next = fat_sectors + directory_sectors + mini_fat_sectors + mini_sectors;
    for ((slot, _), sectors) in builder.large.iter().zip(&large_sectors) {
        if let Some(entry) = builder.slots.get_mut(*slot) {
            entry.start = u32::try_from(next).unwrap_or(0);
        }
        next += sectors;
    }
    if let Some(root) = builder.slots.first_mut() {
        root.child = root_child;
        root.start = u32::try_from(fat_sectors + directory_sectors + mini_fat_sectors).unwrap_or(0);
        root.size = u32::try_from(builder.mini.len()).unwrap_or(0);
    }
    let mut out = header(fat_sectors, directory_sectors, mini_fat_sectors);
    let mut chains = vec![directory_sectors, mini_fat_sectors, mini_sectors];
    chains.extend(&large_sectors);
    out.extend(allocation_table(fat_sectors, &chains));
    out.extend(builder.directory());
    let mut table = vec![FREE; mini_fat_sectors * 128];
    table[..builder.mini_fat.len()].copy_from_slice(&builder.mini_fat);
    out.extend(table.iter().flat_map(|entry| entry.to_le_bytes()));
    let mut mini = builder.mini;
    mini.resize(mini_sectors * 512, 0);
    out.extend_from_slice(&mini);
    for ((_, data), sectors) in builder.large.iter().zip(&large_sectors) {
        let mut padded = data.clone();
        padded.resize(sectors * 512, 0);
        out.extend_from_slice(&padded);
    }
    out
}

/// What a VBA project holds, which the `dir` and `PROJECT` streams may say differently.
struct Spec<'a> {
    /// The name in the `dir` stream.
    name: &'a str,
    /// The name in the `PROJECT` stream.
    project_name: &'a str,
    /// The modules the `dir` stream lists.
    modules: &'a [&'a str],
    /// The modules the `PROJECT` stream lists.
    listed: &'a [&'a str],
    /// The modules that have a stream of their own.
    streams: &'a [&'a str],
    /// Whether the project has references to libraries and controls.
    references: bool,
}

impl<'a> Spec<'a> {
    /// A project whose streams agree.
    fn of(name: &'a str, modules: &'a [&'a str]) -> Self {
        Self {
            name,
            project_name: name,
            modules,
            listed: modules,
            streams: modules,
            references: false,
        }
    }

    fn build(&self) -> Vec<u8> {
        let mut vba = vec![
            Node::Stream(
                "dir".to_owned(),
                compress(&dir_stream(self.name, self.modules, self.references)),
            ),
            Node::Stream(
                "_VBA_PROJECT".to_owned(),
                vec![0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00],
            ),
        ];
        for module in self.streams {
            let source = format!("Attribute VB_Name = \"{module}\"\r\nSub Run()\r\nEnd Sub\r\n");
            vba.push(Node::Stream(
                (*module).to_owned(),
                compress(source.as_bytes()),
            ));
        }
        compound_file(vec![
            Node::Stream(
                "PROJECT".to_owned(),
                project_stream(self.project_name, self.listed),
            ),
            Node::Storage("VBA", vba),
        ])
    }
}

/// A workbook package with the VBA project `project` as its `xl/vbaProject.bin`.
fn macro_package(project: &[u8]) -> Result<Cursor<Vec<u8>>, Box<dyn StdError>> {
    let workbook = workbook_xml(r#"<sheet name="S1" sheetId="1" r:id="rId1"/>"#, "");
    let relationships = rels_xml(&format!(
        r#"<Relationship Id="rId1" Type="{REL_NS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>"#
    ));
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#;
    let root = rels_xml(&format!(
        r#"<Relationship Id="rId1" Type="{REL_NS}/officeDocument" Target="xl/workbook.xml"/>"#
    ));
    let sheet = worksheet_xml(r#"<row r="1"><c r="A1"><v>1</v></c></row>"#);
    // The package holds the project as bytes, which `package` writes as text, so the bytes are
    // written to the package by hand.
    let mut writer = zip::write::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        ("xl/_rels/workbook.xml.rels", relationships.as_bytes()),
        ("xl/worksheets/sheet1.xml", sheet.as_bytes()),
        ("xl/vbaProject.bin", project),
    ] {
        writer.start_file(name, options)?;
        std::io::Write::write_all(&mut writer, bytes)?;
    }
    Ok(writer.finish()?)
}

#[test]
fn the_modules_of_a_vba_project_are_listed_with_the_name_of_the_project()
-> Result<(), Box<dyn StdError>> {
    let project = Spec::of("VBAProject", &["Module1", "Module2"]).build();
    let workbook = sheetrs_xlsx::read(macro_package(&project)?)?;
    let found: Vec<_> = workbook
        .macro_modules()
        .iter()
        .map(|module| (module.kind, module.library.as_str(), module.name.as_str()))
        .collect();
    assert_eq!(
        found,
        [
            (MacroKind::Vba, "VBAProject", "Module1"),
            (MacroKind::Vba, "VBAProject", "Module2"),
        ]
    );
    Ok(())
}

#[test]
fn references_and_a_dir_stream_in_sectors_of_its_own_do_not_hide_the_modules()
-> Result<(), Box<dyn StdError>> {
    let names: Vec<String> = (0..60).map(|number| format!("Module{number}")).collect();
    let modules: Vec<&str> = names.iter().map(String::as_str).collect();
    let project = Spec {
        references: true,
        ..Spec::of("VBAProject", &modules)
    }
    .build();
    let workbook = sheetrs_xlsx::read(macro_package(&project)?)?;
    let found: Vec<&str> = workbook
        .macro_modules()
        .iter()
        .map(|module| module.name.as_str())
        .collect();
    assert_eq!(found, modules);
    Ok(())
}

#[test]
fn a_workbook_without_a_vba_project_holds_no_modules() -> Result<(), Box<dyn StdError>> {
    let workbook = sheetrs_xlsx::read(single_sheet_package_with(
        r#"<row r="1"><c r="A1"><v>1</v></c></row>"#,
        None,
        "",
    )?)?;
    assert_eq!(workbook.macro_modules(), []);
    Ok(())
}

/// VBA projects in the ways a file can lay them out.
fn project_cases() -> Vec<(&'static str, Vec<u8>)> {
    let names: Vec<String> = (0..60).map(|number| format!("Module{number}")).collect();
    let many: Vec<&str> = names.iter().map(String::as_str).collect();
    vec![
        ("one_module", Spec::of("VBAProject", &["Module1"]).build()),
        (
            "two_modules",
            Spec::of("VBAProject", &["Module1", "Module2"]).build(),
        ),
        (
            "another_project_name",
            Spec::of("Budget", &["Module1"]).build(),
        ),
        (
            "document_modules",
            Spec::of("VBAProject", &["ThisWorkbook", "Sheet1"]).build(),
        ),
        ("no_modules", Spec::of("VBAProject", &[]).build()),
        (
            "with_references",
            Spec {
                references: true,
                ..Spec::of("VBAProject", &["Module1"])
            }
            .build(),
        ),
        ("many_modules", Spec::of("VBAProject", &many).build()),
        (
            "name_with_spaces",
            Spec::of("My Project", &["My Module"]).build(),
        ),
        ("empty_project_name", Spec::of("", &["Module1"]).build()),
        (
            "module_only_in_the_dir_stream",
            Spec {
                listed: &["Module1"],
                ..Spec::of("VBAProject", &["Module1", "Module2"])
            }
            .build(),
        ),
        (
            "module_only_in_the_project_stream",
            Spec {
                listed: &["Module1", "Module2"],
                ..Spec::of("VBAProject", &["Module1"])
            }
            .build(),
        ),
        (
            "project_name_differs",
            Spec {
                project_name: "Other",
                ..Spec::of("VBAProject", &["Module1"])
            }
            .build(),
        ),
        (
            "module_without_its_stream",
            Spec {
                streams: &["Module1"],
                ..Spec::of("VBAProject", &["Module1", "Module2"])
            }
            .build(),
        ),
        (
            "stream_without_its_module",
            Spec {
                streams: &["Module1", "Module2"],
                ..Spec::of("VBAProject", &["Module1"])
            }
            .build(),
        ),
        (
            "same_module_twice",
            Spec {
                streams: &["Module1"],
                ..Spec::of("VBAProject", &["Module1", "Module1"])
            }
            .build(),
        ),
        (
            "modules_that_differ_in_case",
            Spec::of("VBAProject", &["Module1", "MODULE1"]).build(),
        ),
        ("not_a_compound_file", b"not a compound file".to_vec()),
    ]
}

#[test]
#[ignore = "needs LibreOffice 26.8 and the Python UNO bridge"]
fn vba_projects_over_xlsx() -> Result<(), Box<dyn StdError>> {
    let cases = project_cases();
    let mut packages = Vec::new();
    for (name, project) in &cases {
        packages.push(Package {
            name: (*name).to_owned(),
            bytes: macro_package(project)?.into_inner(),
            sheet: "S1".to_owned(),
            range: "A1:A1".to_owned(),
        });
    }
    crafted::assert_reads_as_the_oracle(
        DIR,
        "vba",
        "xlsm",
        None,
        &packages,
        |path| Ok(sheetrs_xlsx::read(File::open(path)?)?),
        |_: &CellValue, _: Option<&Cell>| Leniency::Strict,
    )
}
