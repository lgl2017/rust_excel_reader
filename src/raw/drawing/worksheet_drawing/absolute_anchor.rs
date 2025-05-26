use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::{
    excel::XmlReader,
    raw::drawing::{
        graphic::graphic_frame::XlsxGraphicFrame, image::picture::XlsxPicture,
        shape::connection_shape::XlsxConnectionShape,
    },
};

use super::{
    client_data::XlsxClientData, content_part::XlsxContentPart,
    drawing_content_type::XlsxWorksheetDrawingContentType, group_shape::XlsxGroupShape,
    spreadsheet_extent::XlsxSpreadsheetExtent, spreadsheet_position::XlsxSpreadsheetPosition,
    spreadsheet_shape::XlsxShape,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.absoluteanchor?view=openxml-3.0.1
///
/// This element is used as an anchor placeholder for a shape or group of shapes.
/// It anchors the object in the same position relative to sheet position and its extents are in EMU units.
///
/// absoluteAnchor (Absolute Anchor Shape Size)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxAbsoluteAnchor {
    // Child Elements

    // clientData (Client Data)	§20.5.2.3
    pub client_data: Option<XlsxClientData>,

    // contentPart (Content Part)	§20.5.2.12
    pub content_part: Option<XlsxContentPart>,

    // ext (Shape Extent)	§20.5.2.14
    pub extent: Option<XlsxSpreadsheetExtent>,

    // grpSp (Group Shape)	§20.5.2.17
    // pic (Picture)	§20.5.2.25
    // sp (Shape)	§20.5.2.29
    // cxnSp (Connection Shape)	§20.5.2.13
    // graphicFrame (Graphic Frame)
    pub drawing_content: Option<XlsxWorksheetDrawingContentType>,

    // pos (Position)	§20.5.2.26
    pub position: Option<XlsxSpreadsheetPosition>,
}

impl XlsxAbsoluteAnchor {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut anchor = Self {
            client_data: None,
            content_part: None,
            extent: None,
            position: None,
            drawing_content: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clientData" => {
                    anchor.client_data = Some(XlsxClientData::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"contentPart" => {
                    anchor.content_part = Some(XlsxContentPart::load(reader, e)?);
                }
                // load graphic frame first in case there are other type fall back available (ex: pic)
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"graphicFrame" => {
                    anchor.drawing_content = Some(XlsxWorksheetDrawingContentType::GraphicFrame(
                        XlsxGraphicFrame::load(reader, e)?,
                    ));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cxnSp" => {
                    anchor.drawing_content =
                        Some(XlsxWorksheetDrawingContentType::ConnectionShape(
                            XlsxConnectionShape::load(reader, e)?,
                        ));
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ext" => {
                    anchor.extent = Some(XlsxSpreadsheetExtent::load(e)?);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"grpSp" => {
                    anchor.drawing_content = Some(XlsxWorksheetDrawingContentType::GroupShape(
                        XlsxGroupShape::load(reader)?,
                    ))
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pic" => {
                    anchor.drawing_content = Some(XlsxWorksheetDrawingContentType::Picture(
                        XlsxPicture::load(reader, e)?,
                    ))
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sp" => {
                    anchor.drawing_content = Some(XlsxWorksheetDrawingContentType::Shape(
                        XlsxShape::load(reader, e)?,
                    ))
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pos" => {
                    anchor.position = Some(XlsxSpreadsheetPosition::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"absoluteAnchor" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at `absoluteAnchor`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(anchor)
    }
}
