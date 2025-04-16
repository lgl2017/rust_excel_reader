#![allow(dead_code)]
#![allow(unused)]

use std::{
    fmt::format,
    fs::File,
    io::{self, BufReader, Read, Seek, Write},
    path::PathBuf,
    str::FromStr,
};

use anyhow::Context;
use excel_reader::{
    common_types::Coordinate,
    excel::Excel,
    raw::drawing::theme::{self, Theme},
};
use quick_xml::Reader;
use zip::{read::ZipFile, ZipArchive};

/// sample Excels
/// https://create.microsoft.com/en-us/search?query=table&filters=excel
fn main() -> anyhow::Result<()> {
    // let str = "2002-05-30T09:00:00";
    // let d = string_to_datetime(str);
    // println!("{:?}", d);

    // save_excel_as_xmls("custom_styles.xlsx")?;
    let path = "excel_samples/custom_styles.xlsx";
    let reader = BufReader::new(File::open(path)?);
    let mut archive = ZipArchive::new(reader)?;

    let mut excel = Excel::from_path(path)?;
    load_worksheet(excel)?;
    // load_styles(excel)?;
    // let sheets = excel.get_sheets()?;
    // println!("sheets: {:?}", sheets);

    // load_workbook(excel)?;
    // load_strings(excel)?;
    // load_themes(excel)?;
    // let excel = Excel::new(reader)

    // let mut styles = archive.by_name("xl/tables/table1.xml")?;
    // println!("name: {}", styles.name());

    // let mut contents = String::new();
    // styles.read_to_string(&mut contents)?;
    // write_string_to_file(&contents, "table_table.xml")?;

    // let mut styles = archive.by_name("xl/styles.xml")?;
    // println!("name: {}", styles.name());

    // let mut contents = String::new();
    // styles.read_to_string(&mut contents)?;
    // write_string_to_file(&contents, "custom_style_style.xml")?;

    // let mut styles = archive.by_name("xl/sharedStrings.xml")?;
    // println!("name: {}", styles.name());

    // let mut contents = String::new();
    // styles.read_to_string(&mut contents)?;
    // write_string_to_file(&contents, "sample_sharedstring.xml")?;

    // let mut theme = archive.by_name("xl/theme/theme1.xml")?;
    // println!("name: {}", theme.name());

    // let mut contents = String::new();
    // theme.read_to_string(&mut contents)?;
    // write_string_to_file(&contents, "theme_table.xml")?;

    // let mut shared_string = archive.by_name("xl/workbook.xml")?;
    // println!("name: {}", shared_string.name());

    // let mut contents = String::new();
    // shared_string.read_to_string(&mut contents)?;
    // write_string_to_file(&contents, "workbook_table.xml")?;

    // let mut workbook = archive.by_name("xl/workbook.xml")?;
    // println!("name: {}", workbook.name());

    // let mut contents = String::new();
    // workbook.read_to_string(&mut contents)?;
    // write_string_to_file(&contents)?;

    // let mut zip = archive.by_name("xl/worksheets/sheet1.xml")?;
    // println!("name: {}", zip.name());

    // let mut contents = String::new();
    // zip.read_to_string(&mut contents)?;
    // write_string_to_file(&contents, "worksheet_sample.xml")?;

    Ok(())
}

fn save_excel_as_xmls(name: &str) -> anyhow::Result<()> {
    let path = format!("excel_samples/{}", name);
    let reader = BufReader::new(File::open(path)?);
    let mut archive = ZipArchive::new(reader)?;
    for i in 0..archive.len() {
        let mut item = archive.by_index(i)?;

        println!("name: {}", item.name());
        let mut contents = String::new();
        if let Ok(_) = item.read_to_string(&mut contents) {
            let name: Vec<&str> = name.split(".").collect();
            let file_name = format!("{}/{}", name.first().unwrap_or(&"sample"), item.name());
            let _ = write_string_to_file(&contents, &file_name);
        }
    }

    Ok(())
}

fn write_string_to_file(str: &str, name: &str) -> anyhow::Result<()> {
    let path = format!("excel_samples/xmls/{}", name);
    let path = PathBuf::from_str(&path)?;
    let prefix = path.parent().context("failed to find parent")?;
    std::fs::create_dir_all(prefix).context("failed to create parent directory")?;

    let mut file = File::create(path)?;
    file.write_all(str.as_bytes())?;
    file.flush()?;
    Ok(())
}

fn load_worksheet(mut excel: Excel<BufReader<File>>) -> anyhow::Result<()> {
    let sheets = excel.get_sheets()?;
    println!("sheets :{:?}", sheets);
    // excel.get_table_for_worksheet(sheets[0].clone())?;
    let mut worksheet = excel.get_worksheet(&sheets[2].clone())?;
    let tables = worksheet.tables;
    for table in tables {
        println!("table: {}", table.display_name);
        println!("dimension: {:?}", table.dimension);
        println!("columns: {:?}", table.columns);
        println!("table_style name: {:?}", table.table_style.name);
    }
    // println!("dimension: {:?}", worksheet.dimension);
    // println!("merged cells: {:?}", worksheet.merged_cells);
    // let cell = worksheet.get_cell(Coordinate { row: 1, col: 2 });
    // println!("cell: {:?}", cell);
    Ok(())
}

fn load_workbook(mut excel: Excel<BufReader<File>>) -> anyhow::Result<()> {
    let workbook = excel.get_raw_workbook()?.unwrap();
    println!("{:?}", workbook.sheets);
    Ok(())
}

fn load_strings(mut excel: Excel<BufReader<File>>) -> anyhow::Result<()> {
    let string = excel.get_raw_shared_strings()?.unwrap();
    println!("{:?}", string);
    Ok(())
}

fn load_themes(mut excel: Excel<BufReader<File>>) -> anyhow::Result<()> {
    let theme = excel.get_raw_theme()?.unwrap();
    println!("{:?}", theme.theme_elements.clone().unwrap().color_scheme);
    let color_scheme = theme.theme_elements.unwrap().color_scheme.unwrap();

    // println!("lt1: {:?}", color_scheme.get_color(0));
    // println!("dk2: {:?}", color_scheme.get_color(3));

    Ok(())
}

fn load_styles(mut excel: Excel<BufReader<File>>) -> anyhow::Result<()> {
    let themes = excel.get_raw_theme()?.unwrap();
    let color_scheme = themes.theme_elements.unwrap().color_scheme;
    let styles = excel.get_raw_stylesheet()?.unwrap();
    let fonts = styles.clone().fonts.unwrap();
    let font = fonts[0].clone();
    let color = font.color.unwrap();
    println!("font color:{:?}", color);
    // println!("hex:{:?}", color.to_hex(styles.colors, color_scheme));
    // println!(
    //     "styles border: {:?}",
    //     styles.clone().unwrap().borders.unwrap().len()
    // );

    // println!(
    //     "styles fills: {:?}",
    //     styles.clone().unwrap().fills.unwrap().len()
    // );

    // if let Some(colors) = styles.clone().unwrap().colors {
    //     println!(
    //         "styles colors: indexed: {:?}, mru: {:?}",
    //         colors.indexed_colors, colors.mru_colors
    //     );
    // }

    // println!(
    //     "styles fonts: {:?}",
    //     styles.clone().unwrap().fonts.unwrap().len()
    // );

    // println!(
    //     "styles cellStyles: {:?}",
    //     styles.clone().unwrap().cell_styles
    // );

    // println!(
    //     "styles cell_style_xfs: {:?}",
    //     styles.clone().unwrap().cell_style_xfs
    // );

    // println!("styles cell_xfs: {:?}", styles.clone().unwrap().cell_xfs);

    // println!(
    //     "styles differential_xfs: {:?}",
    //     styles.clone().unwrap().differential_xfs
    // );

    // println!(
    //     "styles numbering_formats: {:?}",
    //     styles.clone().unwrap().numbering_formats
    // );

    // if let Some(table_styles) = styles.clone().unwrap().table_styles {
    //     println!(
    //         "styles table_styles: default_pivot_style: {:?}, default_table_style: {:?}, length: {:?}  ",
    //         table_styles.default_pivot_style,
    //         table_styles.default_table_style,
    //         table_styles.table_style.unwrap().len()
    //     );
    // }
    Ok(())
}
