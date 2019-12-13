extern crate xlsx;

use xlsx::workbook::Workbook;

#[test]
fn test() {
    {
        let mut w = Workbook::new("tmp/doc1", "Rust", false);
        w.initialize();

        let number = w.new_format("#,###,###,##0.00");
        let date = w.new_format("dd/mm/yyyy");

        let mut s = w.new_worksheet("Sheet 1", 2);

        s.cell_txt(w.value("Company Name"));
        s.row();

        s.cell_txt(w.value("Title 1"));
        s.cell_txt(w.value("Title 2"));
        s.cell_txt(w.value("Title 3"));
        s.cell_txt(w.value("Title 4"));
        s.row();

        s.cell_num("50.00", number);
        s.cell_num("1300.00", number);
        s.cell_txt(w.value("20"));
        s.cell_fmt(w.value("23/12/2015"), date);
        s.row();
        s.cell_num("50.00", number);
        s.cell_num("1300.00", number);
        s.cell_txt(w.value("20"));
        s.cell_fmt(w.value("23/12/2015"), date);
        s.row();
        s.cell_num("50.00", number);
        s.cell_num("1300.00", number);
        s.cell_txt(w.value("20"));
        s.cell_fmt(w.value("23/12/2015"), date);
        s.row();
        s.cell_num("50.00", number);
        s.cell_num("1300.00", number);
        s.cell_txt(w.value("20"));
        s.cell_fmt(w.value("23/12/2015"), date);
        s.row();
        s.cell_num("50.00", number);
        s.cell_num("1300.00", number);
        s.cell_txt(w.value("20"));
        s.cell_fmt(w.value("23/12/2015"), date);
        s.row();
        s.cell_num("50.00", number);
        s.cell_num("1300.00", number);
        s.cell_txt(w.value("20"));
        s.cell_fmt(w.value("23/12/2015"), date);
        s.row();
        s.cell_num("50.00", number);
        s.cell_num("1300.00", number);
        s.cell_txt(w.value("20"));
        s.cell_fmt(w.value("23/12/2015"), date);
        s.flush();

        s = w.new_worksheet("Sheet 2", 1);
        s.cell_txt(w.value("Other Page"));
        s.row();
        s.cell_num("1", number);
        s.row();
        s.cell_fmt(w.value("12/02/1984"), date);
        s.flush();

        w.flush();

        w.xlsx("tmp/doc1.xlsx");
    }
}
