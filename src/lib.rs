//! Rust Excel utils library

use quick_xml::events::Event;
use quick_xml::reader::Reader;
use std::collections::HashMap;
use std::fs::File;
use std::io::{BufReader, Read};
use zip::ZipArchive;

#[derive(Debug)]
pub struct XlsxWorkbook {
    pub archive: ZipArchive<BufReader<File>>,
    pub sheets: HashMap<String, String>,
    pub defined_names: HashMap<String, String>,
    pub shared_strings: String,
}

impl XlsxWorkbook {
    pub fn open(xls_path: &str) -> Result<XlsxWorkbook, std::io::Error> {
        let f = File::open(xls_path)?;
        let mut archive = ZipArchive::new(BufReader::new(f))?;
        let mut buf_workbook = String::new();

        // Used as internal or external reference to track in XML data
        let mut object_ref = String::new();
        // Used as storage of name or paths of objects in XML data
        let mut object_path = String::new();

        // Mapping relationships. Link rId with paths
        let mut rels: HashMap<String, String> = HashMap::new();
        (archive.by_name("xl/_rels/workbook.xml.rels")?).read_to_string(&mut buf_workbook)?;
        let mut reader = Reader::from_str(&buf_workbook);
        loop {
            match reader.read_event() {
                Err(e) => panic!("Error at position {}: {:?}", reader.error_position(), e),
                Ok(Event::Eof) => break,
                Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"Relationship" {
                        for attr in e.attributes() {
                            match attr {
                                Ok(attr) => match attr.key.as_ref() {
                                    // r:Id
                                    b"Id" => {
                                        object_ref = attr.unescape_value().unwrap().into();
                                    }
                                    b"Target" => {
                                        // Relative path to `xl/`
                                        object_path = attr.unescape_value().unwrap().into();
                                    }
                                    _ => (),
                                },
                                Err(e) => {
                                    panic!("Error at position {}: {:?}", reader.error_position(), e)
                                }
                            }
                        }
                        rels.insert(object_ref.clone(), object_path.clone());
                    }
                }
                _ => (),
            }
        }
        buf_workbook.clear();

        (archive.by_name("xl/workbook.xml")?).read_to_string(&mut buf_workbook)?;
        reader = Reader::from_str(&buf_workbook);
        let mut sheets: HashMap<String, String> = HashMap::new();
        let mut defined_names: HashMap<String, String> = HashMap::new();
        loop {
            match reader.read_event() {
                Err(e) => panic!("Error at position {}: {:?}", reader.error_position(), e),
                Ok(Event::Eof) => break,
                Ok(Event::Empty(e)) => {
                    if e.name().as_ref() == b"sheet" {
                        for attr in e.attributes() {
                            match attr {
                                Ok(attr) => match attr.key.as_ref() {
                                    b"r:id" => {
                                        // Link XML path through r:id using rels
                                        object_path = rels
                                            .get(attr.unescape_value().unwrap().as_ref())
                                            .unwrap()
                                            .into();
                                    }
                                    b"name" => {
                                        // This is the sheet name
                                        object_ref = attr.unescape_value().unwrap().into();
                                    }
                                    _ => (),
                                },
                                Err(e) => {
                                    panic!("Error at position {}: {:?}", reader.error_position(), e)
                                }
                            }
                        }
                        sheets.insert(object_ref.clone(), object_path.clone());
                    }
                }
                Ok(Event::Start(e)) => {
                    if e.name().as_ref() == b"definedName" {
                        for attr in e.attributes() {
                            match attr {
                                Ok(attr) => match attr.key.as_ref() {
                                    b"name" => {
                                        object_path = attr.unescape_value().unwrap().into();
                                    }
                                    _ => (),
                                },
                                Err(e) => {
                                    panic!("Error at position {}: {:?}", reader.error_position(), e)
                                }
                            }
                        }
                        match reader.read_event() {
                            Ok(Event::Text(text)) => {
                                object_ref = text.unescape().unwrap().into();
                            }
                            Err(e) => {
                                panic!("Error at position {}: {:?}", reader.error_position(), e)
                            }
                            _ => (),
                        }
                        defined_names.insert(object_path.clone(), object_ref.clone());
                    }
                }
                _ => (),
            }
        }
        Ok(XlsxWorkbook {
            archive,
            sheets,
            defined_names,
            shared_strings: "".to_string(), // Not implemented yet
        })
    }

    pub fn defined_name_errors(&self) -> HashMap<String, String> {
        let mut def_name_errs: HashMap<String, String> = HashMap::new();
        for (name, range) in &self.defined_names {
            if range.ends_with("#REF!") {
                def_name_errs.insert(name.clone(), range.clone());
            }
        }
        return def_name_errs;
    }
}

pub fn invalid_formulas_by_sheet_path(
    workbook: &mut XlsxWorkbook,
    sheet_path: &str,
) -> Vec<String> {
    let mut buf_workbook = String::new();
    let mut cell = String::new();
    let mut is_error = false;
    let mut cell_errors: Vec<String> = Vec::new();
    (workbook
        .archive
        .by_name(&format!("xl/{sheet_path}"))
        .unwrap())
    .read_to_string(&mut buf_workbook)
    .unwrap();
    let mut reader = Reader::from_str(&buf_workbook);
    loop {
        match reader.read_event() {
            Err(e) => panic!("Error at position {}: {:?}", reader.error_position(), e),
            Ok(Event::Eof) => break,
            Ok(Event::Start(e)) => {
                if e.name().as_ref() == b"c" {
                    for attr in e.attributes() {
                        match attr {
                            Ok(attr) => match attr.key.as_ref() {
                                b"r" => {
                                    cell = attr.unescape_value().unwrap().into();
                                }
                                b"t" => {
                                    is_error = attr.unescape_value().unwrap() == "e";
                                }
                                _ => (),
                            },
                            Err(e) => {
                                panic!("Error at position {}: {:?}", reader.error_position(), e)
                            }
                        }
                    }
                    if is_error {
                        cell_errors.push(cell.clone());
                        is_error = false;
                    }
                }
            }
            _ => (),
        }
    }

    buf_workbook.clear();
    cell_errors
}

pub fn invalid_formulas_all(workbook: &mut XlsxWorkbook) -> HashMap<String, Vec<String>> {
    let mut cells_with_errors_insheet: Vec<String>;
    let mut cells_with_errors: HashMap<String, Vec<String>> = HashMap::new();
    for (sheet_name, sheet_path) in workbook.sheets.clone() {
        cells_with_errors_insheet = invalid_formulas_by_sheet_path(workbook, &sheet_path);

        if cells_with_errors_insheet.len() > 0 {
            cells_with_errors.insert(sheet_name.clone(), cells_with_errors_insheet.clone());
            cells_with_errors_insheet.clear();
        }
    }
    return cells_with_errors;
}
