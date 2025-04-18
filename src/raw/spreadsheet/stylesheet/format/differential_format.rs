use anyhow::bail;
use quick_xml::events::Event;

use crate::{
    excel::XmlReader,
    raw::spreadsheet::stylesheet::{border::XlsxBorder, fill::XlsxFill, font::XlsxFont},
};

use super::{
    alignment::XlsxAlignment, numbering_format::XlsxNumberingFormat, protection::XlsxCellProtection,
};

/// DifferentialFormats: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.differentialformats?view=openxml-3.0.1
///
/// define formatting for all non-cell formatting in this workbook
// tag: dxfs
pub type XlsxDifferentialFormats = Vec<XlsxDifferentialFormat>;

pub(crate) fn load_dxfs(reader: &mut XmlReader) -> anyhow::Result<XlsxDifferentialFormats> {
    let mut buf: Vec<u8> = Vec::new();
    let mut formats: Vec<XlsxDifferentialFormat> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xf" => {
                let format = XlsxDifferentialFormat::load(reader)?;
                formats.push(format);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dxfs" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    Ok(formats)
}

/// DifferentialFormat: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.differentialformat?view=openxml-3.0.1
// tag: dxf
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxDifferentialFormat {
    // children
    pub alignment: Option<XlsxAlignment>,
    pub border: Option<XlsxBorder>,
    pub fill: Option<XlsxFill>,
    pub font: Option<XlsxFont>,
    // tag: numFmt
    pub num_fmt: Option<XlsxNumberingFormat>,
    pub protection: Option<XlsxCellProtection>,
}

impl XlsxDifferentialFormat {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut format = Self {
            alignment: None,
            border: None,
            fill: None,
            font: None,
            num_fmt: None,
            protection: None,
        };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alignment" => {
                    let alignment = XlsxAlignment::load(e)?;
                    format.alignment = Some(alignment)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                    let border = XlsxBorder::load(reader, e)?;
                    format.border = Some(border)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                    let fill = XlsxFill::load(reader)?;
                    format.fill = fill
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                    let font = XlsxFont::load(reader)?;
                    format.font = Some(font)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"numFmt" => {
                    let num_fmt = XlsxNumberingFormat::load(e)?;
                    format.num_fmt = Some(num_fmt)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"protection" => {
                    let protection = XlsxCellProtection::load(e)?;
                    format.protection = Some(protection)
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"dxf" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(format)
    }
}
