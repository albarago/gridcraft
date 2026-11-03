//! A minimal Compound File Binary writer for test fixtures, written from the public [MS-CFB]
//! specification. Test-support only: release builds never contain it.
//!
//! It writes a version 3 file with 512-byte sectors, laid out as: the FAT sectors, then any DIFAT
//! sectors, then the directory, the mini FAT, the mini stream and the regular streams. Every
//! chain is contiguous.
//!
//! - **Mini stream.** Streams shorter than the 4,096-byte cutoff go in the mini stream (64-byte
//!   mini sectors, held in the root entry's chain), as [MS-CFB] §2.6.3 requires and as real
//!   `vbaProject.bin` files do. The header's cutoff is the spec's 0x1000. This is the realistic
//!   layout, so readers are tested on their mini stream path too.
//! - **Directory tree.** Each storage's children form a chain linked through the right sibling,
//!   in [MS-CFB] §2.6.4 order (shorter name first, then by uppercased name), every node black.
//!   That is a valid binary search tree; readers don't rebalance, they only search or walk.
//! - **FAT.** As many FAT sectors as needed. The header DIFAT lists the first 109; beyond that
//!   (files over about 6.8 MiB) DIFAT sectors hold the rest, 127 per sector. Files over
//!   [`MAX_SECTORS`] sectors (128 MiB) are refused with [`CfbError::TooLarge`].
//!
//! Damage hooks ([`Damage`]) let tests write files that a reader must reject:
//!
//! - [`Damage::CyclicChain`]: the stream's first sector points back to itself in the FAT (or the
//!   mini FAT). A reader only meets the loop when it needs the stream's second sector, so the
//!   stream must span at least two sectors (65 bytes in the mini stream, 4,097 in regular
//!   sectors); otherwise writing fails with [`CfbError::CycleNeedsTwoSectors`].
//! - [`Damage::OversizedLength`]: the stream is stored in regular sectors and its directory entry
//!   declares one byte more than those sectors hold (at least the 4,096-byte cutoff, so readers
//!   look for it in regular sectors).
//!
//! [MS-CFB]: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/

/// The file signature ([MS-CFB] §2.2).
pub const SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
/// The most sectors a written file may have (128 MiB of 512-byte sectors).
pub const MAX_SECTORS: usize = 1 << 18;

const SECTOR: usize = 512;
const MINI_SECTOR: usize = 64;
const MINI_CUTOFF: usize = 4096;
const ENTRY: usize = 128;
/// Sector ids per FAT sector.
const PER_FAT: usize = SECTOR / 4;
/// FAT sector ids per DIFAT sector (the last slot links to the next DIFAT sector).
const PER_DIFAT: usize = PER_FAT - 1;
const HEADER_DIFAT: usize = 109;
const MAX_NAME_UNITS: usize = 31;

const DIFSECT: u32 = 0xFFFF_FFFC;
const FATSECT: u32 = 0xFFFF_FFFD;
const ENDOFCHAIN: u32 = 0xFFFF_FFFE;
const FREESECT: u32 = 0xFFFF_FFFF;
const NOSTREAM: u32 = 0xFFFF_FFFF;

const TYPE_STORAGE: u8 = 1;
const TYPE_STREAM: u8 = 2;
const TYPE_ROOT: u8 = 5;
const BLACK: u8 = 1;

/// Damage written into one stream's directory entry or chain.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub enum Damage {
    #[default]
    None,
    /// The stream's first sector links back to itself.
    CyclicChain,
    /// The declared size is larger than the stored data.
    OversizedLength,
}

/// One stream to write: its path from the root (storages, then the stream name) and its bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StreamSpec {
    pub path: Vec<String>,
    pub data: Vec<u8>,
    pub damage: Damage,
}

/// Why a file couldn't be written.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
pub enum CfbError {
    /// A name is empty, longer than 31 UTF-16 units, or holds `/`, `\`, `:` or `!`.
    #[error("compound file name {0:?} is empty, longer than 31 UTF-16 units, or holds / \\ : !")]
    BadName(String),
    /// Two entries in one storage have the same name (ASCII case-insensitive), or a path names
    /// both a stream and a storage.
    #[error("compound file path {0:?} is used twice")]
    Duplicate(String),
    /// A stream path is empty.
    #[error("a compound file stream path is empty")]
    EmptyPath,
    /// [`Damage::CyclicChain`] on a stream that fits in one sector, where no reader would see it.
    #[error("a cyclic chain needs a stream of at least two sectors: {0:?}")]
    CycleNeedsTwoSectors(String),
    /// The file would need more than [`MAX_SECTORS`] sectors.
    #[error("the compound file would be larger than {MAX_SECTORS} sectors")]
    TooLarge,
}

/// One directory entry while building.
struct Entry {
    name: String,
    kind: u8,
    right: u32,
    child: u32,
    start: u32,
    size: u32,
    /// The index into the stream list, for stream entries.
    stream: Option<usize>,
}

impl Entry {
    fn new(name: &str, kind: u8, stream: Option<usize>) -> Self {
        Self { name: name.to_owned(), kind, right: NOSTREAM, child: NOSTREAM, start: ENDOFCHAIN, size: 0, stream }
    }
}

/// Where a stream's data goes.
#[derive(Clone, Copy)]
enum Place {
    Mini { first: usize, count: usize },
    Regular { first: usize, count: usize },
}

/// Writes a compound file holding `streams`. Storages are created from the paths.
///
/// # Errors
///
/// See [`CfbError`].
pub fn write(streams: &[StreamSpec]) -> Result<Vec<u8>, CfbError> {
    let mut entries = vec![Entry::new("Root Entry", TYPE_ROOT, None)];
    add_children(&mut entries, streams, &[], 0)?;

    // Mini stream and regular stream sizes, in sectors, before ids are known.
    let mut mini_sectors = 0usize;
    let mut regular_sectors = 0usize;
    let mut places = Vec::with_capacity(streams.len());
    for s in streams {
        let len = s.data.len();
        let place = if len < MINI_CUTOFF && s.damage != Damage::OversizedLength {
            let count = len.div_ceil(MINI_SECTOR);
            let place = Place::Mini { first: mini_sectors, count };
            mini_sectors = mini_sectors.checked_add(count).ok_or(CfbError::TooLarge)?;
            place
        } else {
            let count = len.div_ceil(SECTOR);
            let place = Place::Regular { first: regular_sectors, count };
            regular_sectors = regular_sectors.checked_add(count).ok_or(CfbError::TooLarge)?;
            place
        };
        if s.damage == Damage::CyclicChain && matches!(place, Place::Mini { count: 0..=1, .. } | Place::Regular { count: 0..=1, .. }) {
            return Err(CfbError::CycleNeedsTwoSectors(s.path.join("/")));
        }
        places.push(place);
    }
    let mini_bytes = mini_sectors.checked_mul(MINI_SECTOR).ok_or(CfbError::TooLarge)?;
    let container_sectors = mini_bytes.div_ceil(SECTOR);
    let mini_fat_sectors = mini_sectors.checked_mul(4).ok_or(CfbError::TooLarge)?.div_ceil(SECTOR);
    let dir_sectors = entries.len().checked_mul(ENTRY).ok_or(CfbError::TooLarge)?.div_ceil(SECTOR);
    let data_sectors = [container_sectors, mini_fat_sectors, regular_sectors].iter().try_fold(dir_sectors, |a, n| a.checked_add(*n));
    let data_sectors = data_sectors.filter(|&n| n <= MAX_SECTORS).ok_or(CfbError::TooLarge)?;
    let (fat_count, difat_count) = fat_size(data_sectors)?;

    // Sector ids, in file order.
    let dir_first = fat_count.checked_add(difat_count).ok_or(CfbError::TooLarge)?;
    let mini_fat_first = dir_first.checked_add(dir_sectors).ok_or(CfbError::TooLarge)?;
    let container_first = mini_fat_first.checked_add(mini_fat_sectors).ok_or(CfbError::TooLarge)?;
    let regular_first = container_first.checked_add(container_sectors).ok_or(CfbError::TooLarge)?;
    let total = regular_first.checked_add(regular_sectors).ok_or(CfbError::TooLarge)?;

    let mut fat = vec![FREESECT; fat_count.checked_mul(PER_FAT).ok_or(CfbError::TooLarge)?];
    for id in 0..fat_count {
        set(&mut fat, id, FATSECT);
    }
    for id in fat_count..dir_first {
        set(&mut fat, id, DIFSECT);
    }
    chain(&mut fat, dir_first, dir_sectors);
    chain(&mut fat, mini_fat_first, mini_fat_sectors);
    chain(&mut fat, container_first, container_sectors);
    let mut mini_fat = vec![FREESECT; mini_fat_sectors.checked_mul(PER_FAT).ok_or(CfbError::TooLarge)?];
    let mut container = vec![0u8; container_sectors.checked_mul(SECTOR).ok_or(CfbError::TooLarge)?];
    let mut regular = vec![0u8; regular_sectors.checked_mul(SECTOR).ok_or(CfbError::TooLarge)?];

    for (index, (s, place)) in streams.iter().zip(&places).enumerate() {
        let (start, table, buffer, offset) = match *place {
            Place::Mini { first, count } => {
                chain(&mut mini_fat, first, count);
                (first, &mut mini_fat, &mut container, first.saturating_mul(MINI_SECTOR))
            }
            Place::Regular { first, count } => {
                let id = regular_first.saturating_add(first);
                chain(&mut fat, id, count);
                (id, &mut fat, &mut regular, first.saturating_mul(SECTOR))
            }
        };
        if let Some(slot) = buffer.get_mut(offset..offset.saturating_add(s.data.len())) {
            slot.copy_from_slice(&s.data);
        }
        if s.damage == Damage::CyclicChain {
            set(table, start, id32(start));
        }
        let size = match (s.damage, *place) {
            (Damage::OversizedLength, Place::Regular { count, .. }) => count.saturating_mul(SECTOR).saturating_add(1).max(MINI_CUTOFF),
            _ => s.data.len(),
        };
        let entry = entries.iter_mut().find(|e| e.stream == Some(index));
        if let Some(e) = entry {
            e.size = u32::try_from(size).map_err(|_| CfbError::TooLarge)?;
            let empty = matches!(place, Place::Mini { count: 0, .. } | Place::Regular { count: 0, .. });
            e.start = if empty { ENDOFCHAIN } else { id32(start) };
        }
    }
    if let Some(root) = entries.first_mut() {
        root.start = if mini_bytes == 0 { ENDOFCHAIN } else { id32(container_first) };
        root.size = u32::try_from(mini_bytes).map_err(|_| CfbError::TooLarge)?;
    }

    let mut out = header(fat_count, difat_count, dir_first, mini_fat_first, mini_fat_sectors);
    out.reserve(total.saturating_mul(SECTOR));
    for v in &fat {
        out.extend_from_slice(&v.to_le_bytes());
    }
    write_difat(&mut out, fat_count, difat_count);
    let dir_start = out.len();
    for e in &entries {
        write_entry(&mut out, e);
    }
    out.resize(dir_start.saturating_add(dir_sectors.saturating_mul(SECTOR)), 0);
    // Free directory entries: type 0 with no siblings or child ([MS-CFB] §2.6.3).
    for slot in out.get_mut(dir_start..).unwrap_or_default().as_chunks_mut::<ENTRY>().0.iter_mut().skip(entries.len()) {
        for at in [68, 72, 76] {
            if let Some(field) = slot.get_mut(at..at + 4) {
                field.copy_from_slice(&NOSTREAM.to_le_bytes());
            }
        }
    }
    for v in &mini_fat {
        out.extend_from_slice(&v.to_le_bytes());
    }
    out.extend_from_slice(&container);
    out.extend_from_slice(&regular);
    Ok(out)
}

/// Adds the children of the storage at `prefix` (entry `parent`), sorted and right-linked, then
/// recurses into child storages. Depth is bounded by the longest path.
fn add_children(entries: &mut Vec<Entry>, streams: &[StreamSpec], prefix: &[String], parent: usize) -> Result<(), CfbError> {
    let under = |p: &[String]| p.len() > prefix.len() && p.iter().zip(prefix).all(|(a, b)| a.eq_ignore_ascii_case(b));
    let mut names: Vec<&str> = Vec::new();
    for s in streams {
        if s.path.is_empty() {
            return Err(CfbError::EmptyPath);
        }
        if let Some(next) = s.path.get(prefix.len()).filter(|_| under(&s.path))
            && !names.iter().any(|n| n.eq_ignore_ascii_case(next))
        {
            check_name(next)?;
            names.push(next);
        }
    }
    names.sort_by_key(|n| order_key(n));
    let mut first = NOSTREAM;
    let mut previous: Option<usize> = None;
    let mut storages = Vec::new();
    for name in names {
        let mut path = prefix.to_vec();
        path.push(name.to_owned());
        let same = |p: &[String]| p.len() == path.len() && p.iter().zip(&path).all(|(a, b)| a.eq_ignore_ascii_case(b));
        let exact: Vec<usize> = streams.iter().enumerate().filter(|(_, s)| same(&s.path)).map(|(i, _)| i).collect();
        let deeper = streams.iter().any(|s| s.path.len() > path.len() && same(s.path.get(..path.len()).unwrap_or_default()));
        let index = entries.len();
        match exact.as_slice() {
            [] => {
                entries.push(Entry::new(name, TYPE_STORAGE, None));
                storages.push((index, path));
            }
            [one] if !deeper => entries.push(Entry::new(name, TYPE_STREAM, Some(*one))),
            _ => return Err(CfbError::Duplicate(path.join("/"))),
        }
        match previous.and_then(|p| entries.get_mut(p)) {
            Some(prev) => prev.right = id32(index),
            None => first = id32(index),
        }
        previous = Some(index);
    }
    if let Some(p) = entries.get_mut(parent) {
        p.child = first;
    }
    for (index, path) in storages {
        add_children(entries, streams, &path, index)?;
    }
    Ok(())
}

/// [MS-CFB] §2.6.4 order: shorter names first, then by uppercased UTF-16.
fn order_key(name: &str) -> (usize, Vec<u16>) {
    (name.encode_utf16().count(), name.to_uppercase().encode_utf16().collect())
}

fn check_name(name: &str) -> Result<(), CfbError> {
    let units = name.encode_utf16().count();
    if units == 0 || units > MAX_NAME_UNITS || name.contains(['/', '\\', ':', '!']) {
        return Err(CfbError::BadName(name.to_owned()));
    }
    Ok(())
}

/// The FAT and DIFAT sector counts for `data` other sectors: the smallest FAT that also maps
/// its own sectors and the DIFAT sectors.
fn fat_size(data: usize) -> Result<(usize, usize), CfbError> {
    let mut fat = 1usize;
    loop {
        let difat = fat.saturating_sub(HEADER_DIFAT).div_ceil(PER_DIFAT);
        let needed = data.saturating_add(fat).saturating_add(difat);
        if fat.saturating_mul(PER_FAT) >= needed {
            return Ok((fat, difat));
        }
        fat = fat.saturating_add(1);
        if fat > MAX_SECTORS.div_ceil(PER_FAT).saturating_add(1) {
            return Err(CfbError::TooLarge);
        }
    }
}

/// Links `count` contiguous sectors from `first` in `table`, ending the chain.
fn chain(table: &mut [u32], first: usize, count: usize) {
    for i in 0..count {
        let id = first.saturating_add(i);
        let next = if i.saturating_add(1) == count { ENDOFCHAIN } else { id32(id.saturating_add(1)) };
        set(table, id, next);
    }
}

fn set(table: &mut [u32], id: usize, value: u32) {
    if let Some(slot) = table.get_mut(id) {
        *slot = value;
    }
}

/// A sector id as `u32`. Ids stay below [`MAX_SECTORS`], so this never saturates in practice.
fn id32(id: usize) -> u32 {
    u32::try_from(id).unwrap_or(ENDOFCHAIN)
}

fn header(fat_count: usize, difat_count: usize, dir_first: usize, mini_fat_first: usize, mini_fat_sectors: usize) -> Vec<u8> {
    let mut h = Vec::with_capacity(SECTOR);
    h.extend_from_slice(&SIGNATURE);
    h.extend_from_slice(&[0; 16]); // CLSID
    h.extend_from_slice(&0x003Eu16.to_le_bytes()); // minor version
    h.extend_from_slice(&3u16.to_le_bytes()); // major version
    h.extend_from_slice(&0xFFFEu16.to_le_bytes()); // byte order
    h.extend_from_slice(&9u16.to_le_bytes()); // sector shift
    h.extend_from_slice(&6u16.to_le_bytes()); // mini sector shift
    h.extend_from_slice(&[0; 6]); // reserved
    h.extend_from_slice(&0u32.to_le_bytes()); // directory sectors: 0 in version 3
    h.extend_from_slice(&id32(fat_count).to_le_bytes());
    h.extend_from_slice(&id32(dir_first).to_le_bytes());
    h.extend_from_slice(&0u32.to_le_bytes()); // transaction signature
    h.extend_from_slice(&id32(MINI_CUTOFF).to_le_bytes());
    let mini_fat_start = if mini_fat_sectors == 0 { ENDOFCHAIN } else { id32(mini_fat_first) };
    h.extend_from_slice(&mini_fat_start.to_le_bytes());
    h.extend_from_slice(&id32(mini_fat_sectors).to_le_bytes());
    let difat_start = if difat_count == 0 { ENDOFCHAIN } else { id32(fat_count) };
    h.extend_from_slice(&difat_start.to_le_bytes());
    h.extend_from_slice(&id32(difat_count).to_le_bytes());
    for i in 0..HEADER_DIFAT {
        let v = if i < fat_count { id32(i) } else { FREESECT };
        h.extend_from_slice(&v.to_le_bytes());
    }
    h
}

/// The DIFAT sectors: FAT sector ids from 109 on, 127 per sector, each linking to the next.
fn write_difat(out: &mut Vec<u8>, fat_count: usize, difat_count: usize) {
    for d in 0..difat_count {
        for slot in 0..PER_DIFAT {
            let fat_id = HEADER_DIFAT.saturating_add(d.saturating_mul(PER_DIFAT)).saturating_add(slot);
            let v = if fat_id < fat_count { id32(fat_id) } else { FREESECT };
            out.extend_from_slice(&v.to_le_bytes());
        }
        let next = if d.saturating_add(1) == difat_count { ENDOFCHAIN } else { id32(fat_count.saturating_add(d).saturating_add(1)) };
        out.extend_from_slice(&next.to_le_bytes());
    }
}

fn write_entry(out: &mut Vec<u8>, e: &Entry) {
    let mut name = [0u8; 64];
    let units: Vec<u16> = e.name.encode_utf16().take(MAX_NAME_UNITS).collect();
    for (slot, unit) in name.as_chunks_mut::<2>().0.iter_mut().zip(&units) {
        slot.copy_from_slice(&unit.to_le_bytes());
    }
    out.extend_from_slice(&name);
    // The name length in bytes, with the terminating null.
    let len = u16::try_from(units.len().saturating_add(1).saturating_mul(2)).unwrap_or(64);
    out.extend_from_slice(&len.to_le_bytes());
    out.push(e.kind);
    out.push(BLACK);
    out.extend_from_slice(&NOSTREAM.to_le_bytes()); // left sibling
    out.extend_from_slice(&e.right.to_le_bytes());
    out.extend_from_slice(&e.child.to_le_bytes());
    out.extend_from_slice(&[0; 16]); // CLSID
    out.extend_from_slice(&[0; 4]); // state bits
    out.extend_from_slice(&[0; 16]); // creation and modified times
    let start = if e.kind == TYPE_STORAGE { 0 } else { e.start };
    out.extend_from_slice(&start.to_le_bytes());
    out.extend_from_slice(&e.size.to_le_bytes());
    out.extend_from_slice(&0u32.to_le_bytes()); // size high bits: 0 in version 3
}
