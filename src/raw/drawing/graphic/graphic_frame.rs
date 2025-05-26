use std::io::Read;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    helper::string_to_bool,
    raw::drawing::{
        non_visual_properties::non_visual_graphic_frame_properties::XlsxNonVisualGraphicFrameProperties,
        shape::transform_2d::XlsxTransform2D,
    },
};

use super::XlsxGraphic;

/// - GraphicFrame: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.graphicframe?view=openxml-3.0.1
/// - SpreadSheet.GraphicFrame: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.graphicframe?view=openxml-3.0.1
///
/// This element describes a single graphical object frame for a spreadsheet which contains a graphical object(Ex: Charts).
/// The graphic object is provided entirely by the document authors who choose to persist this data within the document.
///
/// Example:
/// ```
/// <xdr:graphicFrame macro="">
/// <xdr:nvGraphicFramePr>
/// <xdr:cNvPr id="2" name="Chart 1">
///     <a:extLst>
///         <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
///             <a16:creationId
///                 xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"
///                 id="{9225A8DB-6689-4068-DEE9-84C654EE0AE3}" />
///         </a:ext>
///     </a:extLst>
/// </xdr:cNvPr>
/// <xdr:cNvGraphicFramePr />
/// </xdr:nvGraphicFramePr>
/// <xdr:xfrm>
/// <a:off x="0" y="0" />
/// <a:ext cx="0" cy="0" />
/// </xdr:xfrm>
/// <a:graphic>
/// <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
///     <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
///         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
///         r:id="rId1" />
/// </a:graphicData>
/// </a:graphic>
/// </xdr:graphicFrame>
/// ```
///
/// graphicFrame (Graphic Frame)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxGraphicFrame {
    // Child Elements

    // graphic (Graphic Object)	§20.1.2.2.16
    pub graphic: Option<XlsxGraphic>,

    // nvGraphicFramePr (Non-Visual Properties for a Graphic Frame)	§20.5.2.20
    pub non_visual_graphic_frame_properties: Option<XlsxNonVisualGraphicFrameProperties>,

    // xfrm (2D Transform for Graphic Frames)
    pub transform2d: Option<XlsxTransform2D>,

    // attributes: Only applies to SpreadSheet.GraphicFrame
    // Macro
    // Reference to Custom Function
    // Represents the following attribute in the schema: macro
    pub r#macro: Option<String>,

    // Published
    // Publish to Server Flag
    // Represents the following attribute in the schema: fPublished
    pub published: Option<bool>,
}

impl XlsxGraphicFrame {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            graphic: None,
            non_visual_graphic_frame_properties: None,
            transform2d: None,
            r#macro: None,
            published: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"macro" => {
                            properties.r#macro = Some(string_value);
                        }
                        b"fPublished" => {
                            properties.published = string_to_bool(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"graphic" => {
                    properties.graphic = Some(XlsxGraphic::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"nvGraphicFramePr" => {
                    properties.non_visual_graphic_frame_properties =
                        Some(XlsxNonVisualGraphicFrameProperties::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xfrm" => {
                    properties.transform2d = Some(XlsxTransform2D::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"graphicFrame" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at `graphicFrame`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
