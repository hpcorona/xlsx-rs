extern crate xml_writer;

use std::fs::File;
use std::io::BufWriter;
use std::error::Error;
use std::path::PathBuf;

pub type Xml<'a> = XmlWriter<'a, BufWriter<File>>;

pub use xml_writer::XmlWriter;

pub fn xml_writer_for_file<'a>(path: &PathBuf) -> Result<Xml<'a>, String> {
	match File::create(path) {
		Ok(f) => {
			let buf = BufWriter::new(f);
			let mut xml = XmlWriter::new(buf);
			xml.pretty = false;
			Ok(xml)
		},
		Err(e) => Err(Error::description(&e).to_string())
	}
}
