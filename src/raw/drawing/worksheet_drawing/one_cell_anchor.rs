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
    client_data::XlsxClientData,
    content_part::XlsxContentPart,
    drawing_content_type::XlsxWorksheetDrawingContentType,
    group_shape::XlsxGroupShape,
    marker::{load_from_marker, XlsxFromMarker},
    spreadsheet_extent::XlsxSpreadsheetExtent,
    spreadsheet_shape::XlsxShape,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.onecellanchor?view=openxml-3.0.1
///
/// This element specifies a one cell anchor placeholder for a group, a shape, or a drawing element.
/// It moves with the cell and its extents is in EMU units.
///
/// oneCellAnchor (One Cell Anchor Shape Size)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxOneCellAnchor {
    // Child Elements

    // clientData (Client Data)	§20.5.2.3
    pub client_data: Option<XlsxClientData>,

    // contentPart (Content Part)	§20.5.2.12
    pub content_part: Option<XlsxContentPart>,

    // ext (Shape Extent)	§20.5.2.14
    pub extent: Option<XlsxSpreadsheetExtent>,

    // from (Starting Anchor Point)	§20.5.2.15
    pub from: Option<XlsxFromMarker>,

    // grpSp (Group Shape)	§20.5.2.17
    // pic (Picture)	§20.5.2.25
    // sp (Shape)	§20.5.2.29
    // cxnSp (Connection Shape)	§20.5.2.13
    // graphicFrame (Graphic Frame)
    pub drawing_content: Option<XlsxWorksheetDrawingContentType>,
}

impl XlsxOneCellAnchor {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut anchor = Self {
            client_data: None,
            content_part: None,
            extent: None,
            from: None,
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"from" => {
                    anchor.from = Some(load_from_marker(reader)?);
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

                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"oneCellAnchor" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at `oneCellAnchor`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(anchor)
    }
}
