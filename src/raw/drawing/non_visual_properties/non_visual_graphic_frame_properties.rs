use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::excel::XmlReader;

use super::{
    non_visual_drawing_properties::XlsxNonVisualDrawingProperties,
    non_visual_graphic_frame_drawing_properties::XlsxNonVisualGraphicFrameDrawingProperties,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.nonvisualgraphicframeproperties?view=openxml-3.0.1
///
///
/// This element specifies all non-visual properties for a graphic frame.
/// This element is a container for the non-visual identification properties, shape properties and application properties that are to be associated with a graphic frame.
/// This allows for additional information that does not affect the appearance of the graphic frame to be stored.
///
/// nvGraphicFramePr (Non-Visual Properties for a Graphic Frame)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNonVisualGraphicFrameProperties {
    // Child Elements

    // cNvGraphicFramePr (Non-Visual Graphic Frame Drawing Properties)
    pub non_visual_graphic_frame_drawing_properties:
        Option<XlsxNonVisualGraphicFrameDrawingProperties>,

    // cNvPr (Non-Visual Drawing Properties)
    pub non_visual_drawing_properties: Option<XlsxNonVisualDrawingProperties>,
}

impl XlsxNonVisualGraphicFrameProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            non_visual_graphic_frame_drawing_properties: None,
            non_visual_drawing_properties: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cNvGraphicFramePr" => {
                    properties.non_visual_graphic_frame_drawing_properties =
                        Some(XlsxNonVisualGraphicFrameDrawingProperties::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cNvPr" => {
                    properties.non_visual_drawing_properties =
                        Some(XlsxNonVisualDrawingProperties::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"nvGraphicFramePr" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at `nvGraphicFramePr`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
