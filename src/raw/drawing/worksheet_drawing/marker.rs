use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::{
    excel::XmlReader,
    helper::{extract_text_contents, string_to_int, string_to_unsignedint},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.frommarker?view=openxml-3.0.1
///
/// Example:
///  ```
/// <xdr:from>
///     <xdr:col>2</xdr:col>
///     <xdr:colOff>825500</xdr:colOff>
///     <xdr:row>2</xdr:row>
///     <xdr:rowOff>241300</xdr:rowOff>
/// </xdr:from>
/// ```
pub type XlsxFromMarker = XlsxMarker;

pub(crate) fn load_from_marker(reader: &mut XmlReader<impl Read>) -> anyhow::Result<XlsxToMarker> {
    return XlsxMarker::load(reader, b"from");
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tomarker?view=openxml-3.0.1
///
/// Example:
///  ```
/// <xdr:to>
///     <xdr:col>5</xdr:col>
///     <xdr:colOff>901700</xdr:colOff>
///     <xdr:row>17</xdr:row>
///     <xdr:rowOff>38100</xdr:rowOff>
/// </xdr:to>
/// ```
pub type XlsxToMarker = XlsxMarker;

pub(crate) fn load_to_marker(reader: &mut XmlReader<impl Read>) -> anyhow::Result<XlsxToMarker> {
    return XlsxMarker::load(reader, b"to");
}

#[derive(Debug, Clone, PartialEq)]
pub struct XlsxMarker {
    // child elements:
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.columnid?view=openxml-3.0.1
    ///
    /// 0 based index specifies the column that is used within the from and to elements to specify anchoring information for a shape within a spreadsheet.
    ///
    /// col (Column))
    pub column_id: Option<u64>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.columnoffset?view=openxml-3.0.1
    ///
    /// An ST_Coordinate(simple type represents a one dimensional position or length in EMUs: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/4f890b34-61b8-4d22-beb7-77ac953e66a8) specify the column offset within a cell.
    ///
    /// colOff (Column Offset)
    pub column_offset: Option<i64>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.rowid?view=openxml-3.0.1
    ///
    /// 0 based index specifies specifies the row that is used within the from and to elements to specify anchoring information for a shape within a spreadsheet.
    ///
    /// row (Row)
    pub row_id: Option<u64>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.rowoffset?view=openxml-3.0.1
    ///
    /// An ST_Coordinate(simple type represents a one dimensional position or length in EMUs: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/4f890b34-61b8-4d22-beb7-77ac953e66a8) specify the row offset within a cell.
    ///
    /// rowOff(RowOffset)
    pub row_offset: Option<i64>,
}

impl XlsxMarker {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, tag: &[u8]) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            column_id: None,
            column_offset: None,
            row_id: None,
            row_offset: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"col" => {
                    let str = extract_text_contents(reader, b"col")?;
                    properties.column_id = string_to_unsignedint(&str)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"colOff" => {
                    let str = extract_text_contents(reader, b"colOff")?;
                    properties.column_offset = string_to_int(&str)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"row" => {
                    let str = extract_text_contents(reader, b"row")?;
                    properties.row_id = string_to_unsignedint(&str)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rowOff" => {
                    let str = extract_text_contents(reader, b"rowOff")?;
                    properties.row_offset = string_to_int(&str)
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == tag => break,
                Ok(Event::Eof) => {
                    bail!(
                        "unexpected end of file at Marker: `{}`.",
                        String::from_utf8(tag.to_vec())?
                    )
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
