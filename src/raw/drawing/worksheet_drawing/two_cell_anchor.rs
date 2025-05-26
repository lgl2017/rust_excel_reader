use std::io::Read;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

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
    marker::{load_from_marker, load_to_marker, XlsxFromMarker, XlsxToMarker},
    spreadsheet_shape::XlsxShape,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.twocellanchor?view=openxml-3.0.1
///
/// This element specifies a two cell anchor placeholder for a group, a shape, or a drawing element.
/// It moves with cells and its extents are in EMU(English Metric Unit) units.
///
/// Example:
///
/// ```
/// <xdr:twoCellAnchor editAs="oneCell">
/// <xdr:from>
///     <xdr:col>2</xdr:col>
///     <xdr:colOff>825500</xdr:colOff>
///     <xdr:row>2</xdr:row>
///     <xdr:rowOff>241300</xdr:rowOff>
/// </xdr:from>
/// <xdr:to>
///     <xdr:col>5</xdr:col>
///     <xdr:colOff>901700</xdr:colOff>
///     <xdr:row>17</xdr:row>
///     <xdr:rowOff>38100</xdr:rowOff>
/// </xdr:to>
/// <xdr:pic>
///     <xdr:nvPicPr>
///         <xdr:cNvPr id="2" name="Picture 1">
///             <a:extLst>
///                 <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
///                     <a16:creationId
///                         xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"
///                         id="{747A616F-59EB-8B49-A6F6-BF46B404B319}" />
///                 </a:ext>
///             </a:extLst>
///         </xdr:cNvPr>
///         <xdr:cNvPicPr>
///             <a:picLocks noChangeAspect="1" />
///         </xdr:cNvPicPr>
///     </xdr:nvPicPr>
///     <xdr:blipFill>
///         <a:blip
///             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
///             r:embed="rId1" />
///         <a:stretch>
///             <a:fillRect />
///         </a:stretch>
///     </xdr:blipFill>
///     <xdr:spPr>
///         <a:xfrm>
///             <a:off x="3314700" y="838200" />
///             <a:ext cx="3810000" cy="3581400" />
///         </a:xfrm>
///         <a:prstGeom prst="rect">
///             <a:avLst />
///         </a:prstGeom>
///     </xdr:spPr>
/// </xdr:pic>
/// <xdr:clientData />
/// </xdr:twoCellAnchor>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTwoCellAnchor {
    // Child Elements
    // clientData (Client Data)	§20.5.2.3
    pub client_data: Option<XlsxClientData>,

    // contentPart (Content Part)	§20.5.2.12
    pub content_part: Option<XlsxContentPart>,

    // from (Starting Anchor Point)	§20.5.2.15
    pub from: Option<XlsxFromMarker>,

    // grpSp (Group Shape)	§20.5.2.17
    // pic (Picture)	§20.5.2.25
    // sp (Shape)	§20.5.2.29
    // cxnSp (Connection Shape)	§20.5.2.13
    // graphicFrame (Graphic Frame)
    pub drawing_content: Option<XlsxWorksheetDrawingContentType>,

    // to (Ending Anchor Point)	§20.5.2.32
    pub to: Option<XlsxToMarker>,

    // attributes
    /// editAs (EditAs)
    ///
    /// Positioning and Resizing Behaviors
    ///
    /// Possible Values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.editasvalues?view=openxml-3.0.1
    pub edit_as: Option<String>,
}

impl XlsxTwoCellAnchor {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut anchor = Self {
            client_data: None,
            content_part: None,
            from: None,
            drawing_content: None,
            to: None,
            edit_as: None,
        };

        let attributes = e.attributes();
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"editAs" => {
                            anchor.edit_as = Some(string_value);
                            break;
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"to" => {
                    anchor.to = Some(load_to_marker(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"twoCellAnchor" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at TwoCellAnchor: `twoCellAnchor`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(anchor)
    }
}
