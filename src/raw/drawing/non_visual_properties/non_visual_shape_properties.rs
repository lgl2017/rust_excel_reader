use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::excel::XmlReader;

use super::{
    non_visual_drawing_properties::XlsxNonVisualDrawingProperties,
    non_visual_shape_drawing_properties::XlsxNonVisualShapeDrawingProperties,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.nonvisualshapeproperties?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNonVisualShapeProperties {
    // Child Elements

    // cNvPr (Non-Visual Drawing Properties)
    pub non_visual_drawing_properties: Option<XlsxNonVisualDrawingProperties>,
    // cNvSpPr (Non-Visual Shape Drawing Properties)
    pub non_visual_shape_drawing_properties: Option<XlsxNonVisualShapeDrawingProperties>,
}

impl XlsxNonVisualShapeProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            non_visual_shape_drawing_properties: None,
            non_visual_drawing_properties: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cNvSpPr" => {
                    properties.non_visual_shape_drawing_properties =
                        Some(XlsxNonVisualShapeDrawingProperties::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cNvPr" => {
                    properties.non_visual_drawing_properties =
                        Some(XlsxNonVisualDrawingProperties::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"nvSpPr" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at `nvSpPr`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
