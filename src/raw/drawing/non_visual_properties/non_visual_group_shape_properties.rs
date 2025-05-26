use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::excel::XmlReader;

use super::{
    non_visual_drawing_properties::XlsxNonVisualDrawingProperties,
    non_visual_group_shape_drawing_properties::XlsxNonVisualGroupShapeDrawingProperties,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.nonvisualgroupshapeproperties?view=openxml-3.0.1
///
/// This element specifies all non-visual properties for a group shape.
/// This element is a container for the non-visual identification properties, shape properties and application properties that are to be associated with a group shape.
/// This allows for additional information that does not affect the appearance of the group shape to be stored.
///
/// Example:
/// ```
/// <p:nvGrpSpPr>
///     <p:cNvPr id="10" name="Group 9"/>
///     <p:cNvGrpSpPr/>
///     <p:nvPr/>
/// </p:nvGrpSpPr>
/// ```
///
/// nvGrpSpPr (Non-Visual Properties for a Group Shape)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNonVisualGroupShapeProperties {
    // Child Elements
    // cNvGrpSpPr (Non-Visual Group Shape Drawing Properties)	§20.1.2.2.6
    pub non_visual_group_shape_drawing_properties: Option<XlsxNonVisualGroupShapeDrawingProperties>,

    // cNvPr (Non-Visual Drawing Properties)
    pub non_visual_drawing_properties: Option<XlsxNonVisualDrawingProperties>,
}

impl XlsxNonVisualGroupShapeProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            non_visual_group_shape_drawing_properties: None,
            non_visual_drawing_properties: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cNvGrpSpPr" => {
                    properties.non_visual_group_shape_drawing_properties =
                        Some(XlsxNonVisualGroupShapeDrawingProperties::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cNvPr" => {
                    properties.non_visual_drawing_properties =
                        Some(XlsxNonVisualDrawingProperties::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"nvGrpSpPr" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at `nvGrpSpPr`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
