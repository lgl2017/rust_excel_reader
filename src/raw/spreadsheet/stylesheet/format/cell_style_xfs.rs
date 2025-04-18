use anyhow::bail;
use quick_xml::events::Event;

use crate::excel::XmlReader;

use super::cell_format::XlsxCellFormat;

/// CellStyleFormats: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyleformats?view=openxml-3.0.1
/// Example:
/// ```
/// <cellStyleXfs count="4">
///     <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
///     <xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>
///     <xf numFmtId="0" fontId="3" fillId="0" borderId="1" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/>
///     <xf numFmtId="0" fontId="4" fillId="2" borderId="2" applyNumberFormat="0" applyAlignment="0" applyProtection="0"/>
/// </cellStyleXfs>
/// ```
pub type XlsxCellStyleFormats = Vec<XlsxCellFormat>;

pub(crate) fn load_cell_styles_xfs(reader: &mut XmlReader) -> anyhow::Result<Vec<XlsxCellFormat>> {
    let mut buf: Vec<u8> = Vec::new();
    let mut formats: Vec<XlsxCellFormat> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xf" => {
                let format = XlsxCellFormat::load(reader, e)?;
                formats.push(format);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cellStyleXfs" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    Ok(formats)
}
