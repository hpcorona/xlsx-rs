extern crate zip;

use numberformat::*;
use sharedstrings::*;
use worksheet::*;
use xml::*;

use std::fs;
use std::fs::File;
use std::io::*;
use std::path::PathBuf;

pub struct Workbook<'a> {
    path: PathBuf,
    author: &'a str,
    format: NumberFormat<'a>,
    shared: SharedStrings<'a>,
    worksheets: Xml<'a>,
    sheet_count: u32,
    content: Xml<'a>,
    workbook: Xml<'a>,
    compact: bool,
}

impl<'a> Workbook<'a> {
    pub fn new(directory: &str, author: &'a str, compact: bool) -> Workbook<'a> {
        let path = PathBuf::from(directory);

        {
            // Root
            fs::create_dir_all(path.to_path_buf());

            // _rels
            let mut helper = path.clone();
            helper.push("_rels");
            fs::create_dir_all(helper.to_path_buf());
            helper.pop();

            // docProps
            helper.push("docProps");
            fs::create_dir_all(helper.to_path_buf());
            helper.pop();

            // xl
            helper.push("xl");
            fs::create_dir_all(helper.to_path_buf());

            // xl/_rels
            helper.push("_rels");
            fs::create_dir_all(helper.to_path_buf());
            helper.pop();

            // xl/worksheets
            helper.push("worksheets");
            fs::create_dir_all(helper.to_path_buf());
            helper.pop();

            helper.pop();
        }

        let mut helper = path.clone();
        helper.push("xl");
        helper.push("_rels");
        helper.push("workbook.xml.rels");
        if let Ok(mut xml) = xml_writer_for_file(&helper) {
            xml.dtd("UTF-8");
            xml.begin_elem("Relationships");
            xml.attr_esc(
                "xmlns",
                "http://schemas.openxmlformats.org/package/2006/relationships",
            );

            xml.begin_elem("Relationship");
            xml.attr_esc("Id", "styles");
            xml.attr_esc("Target", "styles.xml");
            xml.attr_esc(
                "Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            );
            xml.end_elem();

            xml.begin_elem("Relationship");
            xml.attr_esc("Id", "sharedstrings");
            xml.attr_esc("Target", "sharedStrings.xml");
            xml.attr_esc(
                "Type",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
            );
            xml.end_elem();

            helper = path.clone();
            helper.push("[Content_Types].xml");
            if let Ok(mut content) = xml_writer_for_file(&helper) {
                content.dtd("UTF-8");
                content.begin_elem("Types");
                content.attr_esc(
                    "xmlns",
                    "http://schemas.openxmlformats.org/package/2006/content-types",
                );

                content.begin_elem("Default");
                content.attr_esc("ContentType", "application/xml");
                content.attr_esc("Extension", "xml");
                content.end_elem();

                content.begin_elem("Default");
                content.attr_esc(
                    "ContentType",
                    "application/vnd.openxmlformats-package.relationships+xml",
                );
                content.attr_esc("Extension", "rels");
                content.end_elem();

                content.begin_elem("Default");
                content.attr_esc("ContentType", "image/jpeg");
                content.attr_esc("Extension", "jpeg");
                content.end_elem();

                content.begin_elem("Override");
                content.attr_esc(
                    "ContentType",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                );
                content.attr_esc("PartName", "/xl/workbook.xml");
                content.end_elem();

                content.begin_elem("Override");
                content.attr_esc(
                    "ContentType",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
                );
                content.attr_esc("PartName", "/xl/styles.xml");
                content.end_elem();

                content.begin_elem("Override");
                content.attr_esc(
                    "ContentType",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
                );
                content.attr_esc("PartName", "/xl/sharedStrings.xml");
                content.end_elem();

                content.begin_elem("Override");
                content.attr_esc(
                    "ContentType",
                    "application/vnd.openxmlformats-package.core-properties+xml",
                );
                content.attr_esc("PartName", "/docProps/core.xml");
                content.end_elem();

                helper = path.clone();
                helper.push("xl");
                helper.push("workbook.xml");
                if let Ok(mut workbook) = xml_writer_for_file(&helper) {
                    workbook.dtd("UTF-8");
                    workbook.begin_elem("workbook");
                    workbook.attr_esc(
                        "xmlns",
                        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                    );
                    workbook.attr_esc(
                        "xmlns:r",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                    );

                    workbook.begin_elem("sheets");

                    let format = NumberFormat::new(&path);
                    let shared = SharedStrings::new(&path);

                    Workbook {
                        path: path,
                        author: author,
                        worksheets: xml,
                        format: format,
                        shared: shared,
                        sheet_count: 0,
                        content: content,
                        workbook: workbook,
                        compact: compact,
                    }
                } else {
                    panic!("cannot create workbook");
                }
            } else {
                panic!("cannot create content xlsx");
            }
        } else {
            panic!("cannot create excel file");
        }
    }

    pub fn initialize(&self) {
        self.write_rels();
        self.write_core();
    }

    fn write_core(&self) {
        let mut helper = self.path.clone();
        helper.push("docProps");
        helper.push("core.xml");
        if let Ok(mut xml) = xml_writer_for_file(&helper) {
            xml.dtd("UTF-8");
            xml.begin_elem("cp:coreProperties");
            xml.attr_esc(
                "xmlns:cp",
                "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            );
            xml.attr_esc("xmlns:dc", "http://purl.org/dc/elements/1.1/");
            xml.attr_esc("xmlns:dcmitype", "http://purl.org/dc/dcmitype/");
            xml.attr_esc("xmlns:dcterms", "http://purl.org/dc/terms/");
            xml.attr_esc("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");

            xml.elem_text("dc:creator", &self.author);

            xml.close();
            xml.flush();
        }
    }

    fn write_rels(&self) {
        let mut helper = self.path.clone();
        helper.push("_rels");
        helper.push(".rels");
        if let Ok(mut xml) = xml_writer_for_file(&helper) {
            xml.dtd("UTF-8");
            xml.begin_elem("Relationships");
            xml.attr_esc(
                "xmlns",
                "http://schemas.openxmlformats.org/package/2006/relationships",
            );

            xml.begin_elem("Relationship");
            xml.attr_esc("Id", "core");
            xml.attr_esc("Target", "docProps/core.xml");
            xml.attr_esc("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
            xml.end_elem();

            xml.begin_elem("Relationship");
            xml.attr_esc("Id", "workbook");
            xml.attr_esc("Target", "xl/workbook.xml");
            xml.attr_esc("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
            xml.end_elem();

            xml.close();
            xml.flush();
        }
    }

    pub fn value(&mut self, value: &'a str) -> u32 {
        if self.compact {
            self.shared.index(value)
        } else {
            self.shared.new_value(value)
        }
    }

    pub fn new_format(&mut self, format: &str) -> u32 {
        self.format.new_format(format)
    }

    pub fn new_worksheet(&mut self, title: &str, headerRows: u32) -> Worksheet<'a> {
        self.sheet_count += 1;
        let idx = self.sheet_count;

        self.worksheets.begin_elem("Relationship");
        self.worksheets.attr_esc("Id", &format!("sheet{}", idx));
        self.worksheets
            .attr_esc("Target", &format!("worksheets/sheet{}.xml", idx));
        self.worksheets.attr_esc(
            "Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        );
        self.worksheets.end_elem();

        self.workbook.begin_elem("sheet");
        self.workbook.attr_esc("name", title);
        self.workbook.attr_esc("r:id", &format!("sheet{}", idx));
        self.workbook.attr_esc("sheetId", &idx.to_string());
        self.workbook.end_elem();

        self.content.begin_elem("Override");
        self.content.attr_esc(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        );
        self.content
            .attr_esc("PartName", &format!("/xl/worksheets/sheet{}.xml", idx));
        self.content.end_elem();

        Worksheet::new(&self.path, idx, headerRows)
    }

    pub fn flush(&mut self) {
        self.shared.flush();
        self.format.flush();

        self.workbook.close();
        self.workbook.flush();

        self.content.close();
        self.content.flush();

        self.worksheets.close();
        self.worksheets.flush();
    }

    pub fn xlsx(&mut self, outfile: &str) {
        let output = PathBuf::from(outfile);
        let file = File::create(output).unwrap();
        let mut zip: zip::ZipWriter<File> = zip::ZipWriter::new(file);

        self.zip_file("_rels/", &mut zip);
        self.zip_file("docProps/", &mut zip);
        self.zip_file("xl/", &mut zip);
        self.zip_file("xl/_rels/", &mut zip);
        self.zip_file("xl/worksheets/", &mut zip);

        self.zip_file("_rels/.rels", &mut zip);

        self.zip_file("docProps/core.xml", &mut zip);

        self.zip_file("xl/_rels/workbook.xml.rels", &mut zip);

        self.zip_file("xl/sharedStrings.xml", &mut zip);

        self.zip_file("xl/styles.xml", &mut zip);

        self.zip_file("xl/workbook.xml", &mut zip);

        self.zip_file("[Content_Types].xml", &mut zip);

        for i in 0..self.sheet_count {
            let idx = i + 1;
            self.zip_file(&format!("xl/worksheets/sheet{}.xml", idx), &mut zip);
        }

        zip.finish();
    }

    fn zip_file(&mut self, file: &str, zip: &mut zip::ZipWriter<File>) {
        {
            let mut options: zip::write::FileOptions = zip::write::FileOptions::default();
            options.compression_method(zip::CompressionMethod::Stored);
            zip.start_file(file, options);
        }

        let mut output = self.path.clone();
        output.push(file);
        if let Ok(mut file) = File::open(output) {
            copy(&mut file, zip);
        } else {
            panic!("cannot copy");
        }
    }

    fn clean(&mut self) {
        let mut helper = self.path.clone();

        helper.push("_rels");
        helper.push(".rels");

        fs::remove_file(&helper);

        helper.pop();

        fs::remove_dir(&helper);

        helper.pop();

        helper.push("docProps");
        helper.push("core.xml");

        fs::remove_file(&helper);

        helper.pop();

        fs::remove_dir(&helper);

        helper.pop();

        helper.push("xl");
        helper.push("_rels");
        helper.push("workbook.xml.rels");

        fs::remove_file(&helper);

        helper.pop();

        fs::remove_dir(&helper);

        helper.pop();

        helper.push("worksheets");

        for i in 0..self.sheet_count {
            let idx = i + 1;
            helper.push(&format!("sheet{}.xml", idx));

            fs::remove_file(&helper);

            helper.pop();
        }

        fs::remove_dir(&helper);

        helper.pop();

        helper.push("sharedStrings.xml");
        fs::remove_file(&helper);
        helper.set_file_name("styles.xml");
        fs::remove_file(&helper);
        helper.set_file_name("workbook.xml");
        fs::remove_file(&helper);
        helper.pop();
        fs::remove_dir(&helper);
        helper.set_file_name("[Content_Types].xml");
        fs::remove_file(&helper);
        helper.pop();
        fs::remove_dir(&helper);
    }
}

impl<'a> Drop for Workbook<'a> {
    fn drop(&mut self) {
        self.clean();
    }
}
