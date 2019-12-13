use xml::*;

use std::path::PathBuf;
use std::vec::Vec;

pub struct NumberFormat<'a> {
    data: Vec<u32>,
    writer: Xml<'a>,
    next_format: u32,
}

impl<'a> NumberFormat<'a> {
    pub fn new(path: &PathBuf) -> NumberFormat<'a> {
        let mut shared_path = path.clone();
        shared_path.push("xl");
        shared_path.push("styles.xml");
        if let Ok(mut xml) = xml_writer_for_file(&shared_path) {
            xml.dtd("UTF-8");
            xml.begin_elem("styleSheet");
            xml.attr_esc("mc:Ignorable", "x14ac");
            xml.attr_esc(
                "xmlns",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            );
            xml.attr_esc(
                "xmlns:mc",
                "http://schemas.openxmlformats.org/markup-compatibility/2006",
            );
            xml.attr_esc(
                "xmlns:x14ac",
                "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
            );

            xml.begin_elem("numFmts");

            NumberFormat {
                data: Vec::new(),
                writer: xml,
                next_format: 164,
            }
        } else {
            panic!("cannot create styles")
        }
    }

    pub fn new_format(&mut self, format: &str) -> u32 {
        let current = self.next_format;
        self.writer.begin_elem("numFmt");
        self.writer.attr_esc("formatCode", format);
        self.writer.attr_esc("numFmtId", &current.to_string());
        self.writer.end_elem();

        self.data.push(current);

        self.next_format += 1;

        self.data.len() as u32
    }

    pub fn flush(&mut self) {
        self.writer.end_elem();

        self.writer.begin_elem("fonts");
        self.writer.attr_esc("x14ac:knownFonts", "1");
        self.writer.begin_elem("font");
        self.writer.begin_elem("sz");
        self.writer.attr_esc("val", "12");
        self.writer.end_elem();
        self.writer.begin_elem("name");
        self.writer.attr_esc("val", "Calibri");
        self.writer.end_elem();
        self.writer.begin_elem("family");
        self.writer.attr_esc("val", "2");
        self.writer.end_elem();
        self.writer.end_elem();
        self.writer.end_elem();

        self.writer.begin_elem("fills");
        self.writer.begin_elem("fill");
        self.writer.begin_elem("patternFill");
        self.writer.attr_esc("patternType", "none");
        self.writer.end_elem();
        self.writer.end_elem();
        self.writer.end_elem();

        self.writer.begin_elem("borders");
        self.writer.begin_elem("border");
        self.writer.begin_elem("left");
        self.writer.end_elem();
        self.writer.begin_elem("right");
        self.writer.end_elem();
        self.writer.begin_elem("top");
        self.writer.end_elem();
        self.writer.begin_elem("bottom");
        self.writer.end_elem();
        self.writer.begin_elem("diagonal");
        self.writer.end_elem();
        self.writer.end_elem();
        self.writer.end_elem();

        self.writer.begin_elem("cellXfs");
        self.writer.begin_elem("xf");
        self.writer.attr_esc("applyNumberFormat", "1");
        self.writer.attr_esc("numFmtId", "49");
        self.writer.end_elem();

        for id in self.data.iter() {
            self.writer.begin_elem("xf");
            self.writer.attr_esc("applyNumberFormat", "1");
            self.writer.attr_esc("numFmtId", &id.to_string());
            self.writer.end_elem();
        }
        self.writer.end_elem();

        self.writer.close();
        self.writer.flush();
    }
}
