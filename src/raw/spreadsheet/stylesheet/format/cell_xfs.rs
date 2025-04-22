use anyhow::bail;
use quick_xml::events::Event;
use std::io::Read;

use crate::excel::XmlReader;

use super::cell_format::XlsxCellFormat;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformats?view=openxml-3.0.1
/// Example:
/// ```
/// <cellXfs count="1">
///     <xf numFmtId="0" fontId="0" applyNumberFormat="0" applyFont="1" applyFill="0" applyBorder="0" applyAlignment="1" applyProtection="0">
///         <alignment vertical="top" wrapText="1" />
///     </xf>
/// </cellXfs>
/// ```
pub type XlsxCellFormats = Vec<XlsxCellFormat>;

pub(crate) fn load_cell_xfs(reader: &mut XmlReader<impl Read>) -> anyhow::Result<XlsxCellFormats> {
    let mut buf: Vec<u8> = Vec::new();
    let mut formats: Vec<XlsxCellFormat> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xf" => {
                let format = XlsxCellFormat::load(reader, e)?;
                formats.push(format);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cellXfs" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    Ok(formats)
}
