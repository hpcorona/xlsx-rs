
use std::path::PathBuf;
use std::fmt;

use xml::*;
use sharedstrings::*;

pub struct Worksheet<'a> {
	writer: Xml<'a>
}

impl<'a> Worksheet<'a> {

	pub fn new(path: &PathBuf, idx: u32, headerRows: u32) -> Worksheet<'a> {
		let mut filePath = path.clone();
		filePath.push("xl");
		filePath.push("worksheets");
		filePath.push(format!("sheet{}.xml", idx));

		if let Ok(mut xml) = xml_writer_for_file(&filePath) {
			xml.dtd("UTF-8");
			xml.begin_elem("worksheet");
				xml.attr_esc("mc:Ignorable", "x14ac");
				xml.attr_esc("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
				xml.attr_esc("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
				xml.attr_esc("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
				xml.attr_esc("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

				xml.begin_elem("sheetViews");
					xml.begin_elem("sheetView");
						xml.attr_esc("workbookViewId", "0");

			if headerRows > 0 {
				let rows = headerRows + 1;
						xml.begin_elem("pane");
							xml.attr_esc("activePane", "bottomLeft");
							xml.attr_esc("state", "frozen");
							xml.attr_esc("topLeftCell", &format!("A{}", rows));
							xml.attr_esc("ySplit", &headerRows.to_string());
						xml.end_elem();
			}
					xml.end_elem();
				xml.end_elem();

				xml.begin_elem("sheetFormatPr");
					xml.attr_esc("baseColWidth", "10");
					xml.attr_esc("defaultRowHeight", "15");
					xml.attr_esc("x14ac:dyDescent", "0");
				xml.end_elem();

				xml.begin_elem("sheetData");
					xml.begin_elem("row");

			Worksheet {
				writer: xml
			}
		} else {
			panic!("cannot create worksheet")
		}
	}

	pub fn cell_txt(&mut self, value: u32) {
		self.writer.begin_elem("c");
			self.writer.attr_esc("s", "0");
			self.writer.attr_esc("t", "s");
			self.writer.elem_text("v", &value.to_string());
		self.writer.end_elem();
	}

	pub fn cell_num(&mut self, value: &str, style: u32) {
		self.writer.begin_elem("c");
			self.writer.attr_esc("s", &style.to_string());
			self.writer.elem_text("v", value);
		self.writer.end_elem();
	}

	pub fn cell_fmt(&mut self, value: u32, style: u32) {
		self.writer.begin_elem("c");
			self.writer.attr_esc("s", &style.to_string());
			self.writer.attr_esc("t", "s");
			self.writer.elem_text("v", &value.to_string());
		self.writer.end_elem();
	}

	pub fn row(&mut self) {
		self.writer.end_elem();
		self.writer.begin_elem("row");
	}

	pub fn flush(&mut self) {
		self.writer.close();
		self.writer.flush();
	}
}
