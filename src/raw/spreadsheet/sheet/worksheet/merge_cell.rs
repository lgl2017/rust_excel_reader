use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    common_types::{Coordinate, Dimension},
    excel::XmlReader,
    helper::a1_dimension_to_row_col,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.mergecells?view=openxml-3.0.1
///
/// This collection expresses all the merged cells in the sheet.
///
/// Example:
/// ```
/// <mergeCells>
///   <mergeCell ref="C2:F2"/>
///   <mergeCell ref="B19:C20"/>
///   <mergeCell ref="E19:G19"/>
/// </mergeCells>
/// ```
pub type MergeCells = Vec<MergeCell>;

pub(crate) fn load_merge_cells(reader: &mut XmlReader) -> anyhow::Result<MergeCells> {
    let mut cells: MergeCells = vec![];

    let mut buf = Vec::new();
    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"mergeCell" => {
                if let Some(cell) = load_merge_cell(e)? {
                    cells.push(cell);
                }
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"mergeCells" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(cells)
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.mergecell?view=openxml-3.0.1
///
/// A single merged cell
///
/// Example
/// ```
/// <mergeCell ref="A1:B1" />
/// ```
pub type MergeCell = Dimension;
pub(crate) fn load_merge_cell(e: &BytesStart) -> anyhow::Result<Option<MergeCell>> {
    let attributes = e.attributes();

    for a in attributes {
        match a {
            Ok(a) => match a.key.local_name().as_ref() {
                b"ref" => {
                    let value = a.value.as_ref();
                    let dimension = a1_dimension_to_row_col(value)?;
                    return Ok(Some(MergeCell {
                        start: Coordinate::from_point(dimension.0),
                        end: Coordinate::from_point(dimension.1),
                    }));
                }
                _ => {}
            },
            Err(error) => {
                bail!(error.to_string())
            }
        }
    }
    Ok(None)
}
