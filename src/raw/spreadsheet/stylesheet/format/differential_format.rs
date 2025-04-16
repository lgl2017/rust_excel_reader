use anyhow::bail;
use quick_xml::events::Event;

use crate::{
    excel::XmlReader,
    raw::spreadsheet::stylesheet::{border::Border, fill::Fill, font::Font},
};

use super::{alignment::Alignment, numbering_format::NumberingFormat, protection::Protection};

/// DifferentialFormats: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.differentialformats?view=openxml-3.0.1
///
/// define formatting for all non-cell formatting in this workbook
// tag: dxfs
pub type DifferentialFormats = Vec<DifferentialFormat>;

pub(crate) fn load_dxfs(reader: &mut XmlReader) -> anyhow::Result<DifferentialFormats> {
    let mut buf: Vec<u8> = Vec::new();
    let mut formats: Vec<DifferentialFormat> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xf" => {
                let format = DifferentialFormat::load(reader)?;
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
pub struct DifferentialFormat {
    // children
    pub alignment: Option<Alignment>,
    pub border: Option<Border>,
    pub fill: Option<Fill>,
    pub font: Option<Font>,
    // tag: numFmt
    pub num_fmt: Option<NumberingFormat>,
    pub protection: Option<Protection>,
}

impl DifferentialFormat {
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
                    let alignment = Alignment::load(e)?;
                    format.alignment = Some(alignment)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"border" => {
                    let border = Border::load(reader, e)?;
                    format.border = Some(border)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                    let fill = Fill::load(reader)?;
                    format.fill = fill
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                    let font = Font::load(reader)?;
                    format.font = Some(font)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"numFmt" => {
                    let num_fmt = NumberingFormat::load(e)?;
                    format.num_fmt = Some(num_fmt)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"protection" => {
                    let protection = Protection::load(e)?;
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
