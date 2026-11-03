//! [`MemSource`]: an in-memory [`StreamSource`] for tests and the Fixture_Builder.

use std::collections::{BTreeMap, BTreeSet};

use super::StreamSource;

/// Streams held in memory, keyed by path. Names match ASCII case-insensitively, like a compound
/// file. Storages exist implicitly: `children(path)` lists the distinct next names under every
/// stream path that starts with `path`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct MemSource {
    streams: BTreeMap<Vec<String>, Vec<u8>>,
    /// Paths whose `stream` (or `children`) call returns `Err`.
    failing: BTreeSet<Vec<String>>,
}

impl MemSource {
    /// An empty source.
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// Adds (or replaces) the stream at `path`.
    #[must_use]
    pub fn with(mut self, path: &[&str], bytes: impl Into<Vec<u8>>) -> Self {
        self.insert(path, bytes);
        self
    }

    /// Adds (or replaces) the stream at `path`.
    pub fn insert(&mut self, path: &[&str], bytes: impl Into<Vec<u8>>) {
        let key = owned(path);
        self.streams.retain(|k, _| !same(k, path));
        self.streams.insert(key, bytes.into());
    }

    /// Removes the stream at `path`, if any.
    pub fn remove(&mut self, path: &[&str]) {
        self.streams.retain(|k, _| !same(k, path));
    }

    /// Makes `stream(path)` and `children(path)` return `Err`.
    #[must_use]
    pub fn failing(mut self, path: &[&str]) -> Self {
        self.failing.insert(owned(path));
        self
    }

    fn fails(&self, path: &[&str]) -> bool {
        self.failing.iter().any(|k| same(k, path))
    }
}

impl StreamSource for MemSource {
    fn stream(&self, path: &[&str]) -> Result<Option<Vec<u8>>, String> {
        if self.fails(path) {
            return Err(format!("test read error at {}", path.join("/")));
        }
        Ok(self.streams.iter().find(|(k, _)| same(k, path)).map(|(_, v)| v.clone()))
    }

    fn children(&self, path: &[&str]) -> Result<Vec<String>, String> {
        if self.fails(path) {
            return Err(format!("test list error at {}", path.join("/")));
        }
        let mut names: Vec<String> = Vec::new();
        for key in self.streams.keys() {
            let under = key.len() > path.len() && key.iter().zip(path).all(|(k, p)| k.eq_ignore_ascii_case(p));
            if let Some(next) = key.get(path.len()).filter(|_| under)
                && !names.iter().any(|n| n.eq_ignore_ascii_case(next))
            {
                names.push(next.clone());
            }
        }
        Ok(names)
    }
}

fn owned(path: &[&str]) -> Vec<String> {
    path.iter().map(|s| (*s).to_owned()).collect()
}

fn same(key: &[String], path: &[&str]) -> bool {
    key.len() == path.len() && key.iter().zip(path).all(|(k, p)| k.eq_ignore_ascii_case(p))
}
