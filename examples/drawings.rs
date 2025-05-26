use excel_reader::excel::Excel;

fn main() -> anyhow::Result<()> {
    let path = "examples/workbook_drawing.xlsx";
    let mut excel = Excel::from_path(path)?;

    let sheets = excel.get_sheets()?;
    let worksheet = excel.get_worksheet(&sheets[0].clone())?;
    let drawings = worksheet.get_drawings();
    println!("drawings: {}", drawings.len());

    for (_, drawing) in drawings.into_iter().enumerate() {
        println!("---------------");
        println!(
            "anchor: {}",
            serde_json::to_string_pretty(&drawing.anchor.clone())?
        );
        println!(
            "content: {}",
            serde_json::to_string_pretty(&drawing.content.clone())?
        );
    }

    Ok(())
}
