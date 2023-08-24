#[macro_use] extern crate serde_derive;

pub mod reader;

pub use crate::reader::parse_xlsx;

#[cfg(test)]
mod tests;
