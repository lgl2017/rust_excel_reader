use std::io::Read;

use anyhow::bail;
use graphic_data::XlsxGraphicData;
use quick_xml::events::Event;

use crate::excel::XmlReader;

pub mod graphic_data;
pub mod graphic_frame;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.graphic?view=openxml-3.0.1
///
/// This element specifies the existence of a single graphic object.
/// Document authors should refer to this element when they wish to persist a graphical object of some kind.
/// The specification for this graphical object is provided entirely by the document author and referenced within the graphicData child element.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxGraphic {
    // Child Elements	Subclause
    // graphicData (Graphic Object Data)
    pub graphic_data: Option<XlsxGraphicData>,
}

impl XlsxGraphic {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut graphic = Self { graphic_data: None };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"graphicData" => {
                    graphic.graphic_data = Some(XlsxGraphicData::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"graphic" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at Graphic: `graphic`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(graphic)
    }
}
