extern crate zip;
extern crate serde_xml;

pub mod reader;

pub use reader::parse_xlsx;

#[cfg(test)]
mod tests;
