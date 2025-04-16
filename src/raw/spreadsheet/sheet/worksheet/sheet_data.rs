use anyhow::bail;
use quick_xml::events::Event;

use crate::excel::XmlReader;

use super::row::Row;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetdata?view=openxml-3.0.1
///
/// This collection represents the cell table itself.
/// This collection expresses information about each cell, grouped together by rows in the worksheet.
///
/// Example:
/// ```
/// <sheetData>
///     <row r="1" ht="27.65" customHeight="1">
///         <c r="A1" t="s" s="2">
///             <v>0</v>
///         </c>
///         <c r="B1" s="3" />
///         <c r="C1" s="4" />
///         <c r="D1" s="4" />
///         <c r="E1" s="5" />
///     </row>
///     <row r="2" ht="20.25" customHeight="1">
///         <c r="A2" t="s" s="6">
///             <v>1</v>
///         </c>
///         <c r="B2" t="s" s="6">
///             <v>2</v>
///         </c>
///         <c r="C2" s="7" />
///         <c r="D2" s="8" />
///         <c r="E2" s="9" />
///     </row>
/// </sheetData>
/// ```
///
/// sheetData (Sheet Data)
#[derive(Debug, Clone, PartialEq)]
pub struct SheetData {
    // Child Elements
    /// row (Row)
    pub rows: Option<Vec<Row>>,
}

impl SheetData {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut rows: Vec<Row> = vec![];

        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"row" => {
                    rows.push(Row::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sheetData" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `row`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        return Ok(Self { rows: Some(rows) });
    }
}
