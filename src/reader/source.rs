//! Where a workbook comes from (a file or bytes) and which reader opens it.
//!
//! The format is detected from the content, not the file name: ZIP packages
//! are told apart by their parts (`xl/workbook.bin` = xlsb, `content.xml` =
//! ods, otherwise xlsx/xlsm) and OLE2 compound files are xls. Everything here
//! runs without the GIL, so errors are plain [`Failure`] values that are
//! turned into Python exceptions by the caller.

use std::fs::File;
use std::io::{self, BufReader, Cursor, Read, Seek, SeekFrom};
use std::path::{Path, PathBuf};

use calamine::{Ods, Reader, Sheets, Xls, Xlsb, Xlsx, XlsxError};
use pyo3::PyErr;

use crate::common::bytes::SharedBytes;
use crate::common::errors::{file_err, read_err};

/// A reader over a file on disk or over bytes held in memory (shared, so a
/// second reader can be made for the same data without copying it).
pub enum Source {
    File(BufReader<File>),
    Memory(Cursor<SharedBytes>),
}

/// How to make a fresh [`Source`] for the same workbook (used for the few
/// parts calamine does not expose, such as table definitions).
#[derive(Clone)]
pub enum Origin {
    Path(PathBuf),
    Bytes(SharedBytes),
}

impl Origin {
    pub fn open(&self) -> Result<Source, Failure> {
        match self {
            Origin::Path(path) => open_file(path),
            Origin::Bytes(bytes) => Ok(Source::Memory(Cursor::new(bytes.clone()))),
        }
    }
}

impl Read for Source {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        match self {
            Source::File(f) => f.read(buf),
            Source::Memory(c) => c.read(buf),
        }
    }
}

impl Seek for Source {
    fn seek(&mut self, pos: SeekFrom) -> io::Result<u64> {
        match self {
            Source::File(f) => f.seek(pos),
            Source::Memory(c) => c.seek(pos),
        }
    }
}

/// An error raised while opening or reading, before it becomes a Python exception.
#[derive(Debug)]
pub enum Failure {
    /// The file could not be opened or read (FileError).
    File(String),
    /// The content could not be parsed (ReadError).
    Read(String),
}

impl Failure {
    pub fn into_py(self) -> PyErr {
        match self {
            Failure::File(message) => file_err(message),
            Failure::Read(message) => read_err(message),
        }
    }
}

impl From<Failure> for PyErr {
    fn from(value: Failure) -> Self {
        value.into_py()
    }
}

/// The spreadsheet formats calamine reads.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FileKind {
    Xlsx,
    Xlsb,
    Xls,
    Ods,
}

impl FileKind {
    pub fn name(self) -> &'static str {
        match self {
            FileKind::Xlsx => "xlsx",
            FileKind::Xlsb => "xlsb",
            FileKind::Xls => "xls",
            FileKind::Ods => "ods",
        }
    }
}

pub fn open_file(path: &Path) -> Result<Source, Failure> {
    let file = File::open(path)
        .map_err(|e| Failure::File(format!("cannot open {}: {e}", path.display())))?;
    if file.metadata().is_ok_and(|m| m.is_dir()) {
        return Err(Failure::File(format!(
            "cannot open {}: it is a directory",
            path.display()
        )));
    }
    Ok(Source::File(BufReader::with_capacity(64 * 1024, file)))
}

const ZIP_MAGIC: &[u8] = b"PK";
const CFB_MAGIC: &[u8] = &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const FORMATS: &str = "expected an xlsx, xlsm, xlsb, xls or ods spreadsheet";

fn io_failure(e: io::Error) -> Failure {
    Failure::File(format!("cannot read the workbook: {e}"))
}

/// Work out the format from the first bytes (and, for ZIP packages, the part names).
pub fn detect(source: &mut Source) -> Result<FileKind, Failure> {
    let mut head = [0u8; 8];
    let mut filled = 0;
    while filled < head.len() {
        match source.read(&mut head[filled..]) {
            Ok(0) => break,
            Ok(n) => filled += n,
            Err(e) if e.kind() == io::ErrorKind::Interrupted => continue,
            Err(e) => return Err(io_failure(e)),
        }
    }
    source.seek(SeekFrom::Start(0)).map_err(io_failure)?;
    let head = &head[..filled];

    if head.is_empty() {
        return Err(Failure::Read(format!("the file is empty: {FORMATS}")));
    }
    if head.starts_with(CFB_MAGIC) {
        return Ok(FileKind::Xls);
    }
    if !head.starts_with(ZIP_MAGIC) {
        return Err(Failure::Read(format!(
            "not a spreadsheet (unrecognised content): {FORMATS}"
        )));
    }

    let kind = {
        let archive = zip::ZipArchive::new(&mut *source).map_err(|e| {
            Failure::Read(format!(
                "not a spreadsheet (damaged or unsupported ZIP package: {e}): {FORMATS}"
            ))
        })?;
        let mut has_xlsb = false;
        let mut has_ods = false;
        let mut has_ooxml = false;
        for name in archive.file_names().flatten() {
            let name = name.to_ascii_lowercase();
            if name.ends_with("workbook.bin") {
                has_xlsb = true;
            } else if name == "content.xml" {
                has_ods = true;
            } else if name == "[content_types].xml" || name.ends_with("workbook.xml") {
                has_ooxml = true;
            }
        }
        if has_xlsb {
            FileKind::Xlsb
        } else if has_ooxml {
            FileKind::Xlsx
        } else if has_ods {
            FileKind::Ods
        } else {
            return Err(Failure::Read(format!(
                "not a spreadsheet (a ZIP package without a workbook): {FORMATS}"
            )));
        }
    };
    source.seek(SeekFrom::Start(0)).map_err(io_failure)?;
    Ok(kind)
}

fn parse_failure(kind: FileKind, error: impl std::fmt::Display) -> Failure {
    Failure::Read(format!("cannot read the {} workbook: {error}", kind.name()))
}

/// Open `source` with the reader for its format.
pub fn open(mut source: Source, origin: &Origin) -> Result<(FileKind, Sheets<Source>), Failure> {
    let kind = detect(&mut source)?;
    let sheets = match kind {
        FileKind::Xlsx => match Xlsx::new(source) {
            Ok(r) => Sheets::Xlsx(r),
            Err(XlsxError::Password) => return Err(password_failure()),
            Err(XlsxError::Io(e)) => return Err(io_failure(e)),
            Err(e) => return Err(parse_failure(kind, e)),
        },
        FileKind::Xlsb => Sheets::Xlsb(Xlsb::new(source).map_err(|e| parse_failure(kind, e))?),
        FileKind::Ods => Sheets::Ods(Ods::new(source).map_err(|e| parse_failure(kind, e))?),
        FileKind::Xls => match Xls::new(source) {
            Ok(r) => Sheets::Xls(r),
            Err(error) => {
                // An encrypted xlsx is also an OLE2 compound file.
                if let Ok(fresh) = origin.open()
                    && let Err(XlsxError::Password) = Xlsx::new(fresh)
                {
                    return Err(password_failure());
                }
                return Err(parse_failure(kind, error));
            }
        },
    };
    Ok((kind, sheets))
}

fn password_failure() -> Failure {
    Failure::Read(
        "the workbook is password-protected (encrypted); save an unencrypted copy first".into(),
    )
}
