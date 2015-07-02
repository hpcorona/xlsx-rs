
use xml::*;

use std::path::PathBuf;
use std::fs::File;
use std::io::BufWriter;
use std::collections::BTreeMap;

pub struct SharedStrings<'a> {
	data: BTreeMap<&'a str, u32>,
	writer: Xml<'a>
}

impl<'a> SharedStrings<'a> {
	pub fn new(path: &PathBuf) -> SharedStrings<'a> {
		let mut sharedPath = path.clone();
		sharedPath.push("xl");
		sharedPath.push("sharedStrings.xml");
		if let Ok(mut xml) = xml_writer_for_file(&sharedPath) {
			xml.dtd("UTF-8");
			xml.begin_elem("sst");
				xml.attr_esc("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
				xml.attr_esc("count", "0");
				xml.attr_esc("uniqueCount", "0");

			SharedStrings {
				data: BTreeMap::new(),
				writer: xml
			}
		} else {
			panic!("cannot create shared strings")
		}
	}

	pub fn index(&mut self, value: &'a str) -> u32 {
		{
			let result: Option<&u32> = self.data.get(value);
			if let Some(&val) = result {
				return val;
			}
		}

		let current = self.data.len() as u32;
		self.data.insert(value, current);

		self.writer.begin_elem("si");
		self.writer.elem_text("t", value);
		self.writer.end_elem();

		current
	}

	pub fn flush(&mut self) {
		self.writer.close();
		self.writer.flush();
	}
}
