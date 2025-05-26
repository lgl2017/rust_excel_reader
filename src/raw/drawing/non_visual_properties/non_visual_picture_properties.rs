use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::excel::XmlReader;

use super::{
    non_visual_drawing_properties::XlsxNonVisualDrawingProperties,
    non_visual_picture_drawing_properties::XlsxNonVisualPictureDrawingProperties,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.nonvisualpictureproperties?view=openxml-3.0.1
///
/// This element specifies all non-visual properties for a picture.
/// This element is a container for the non-visual identification properties, shape properties and application properties that are to be associated with a picture.
/// This allows for additional information that does not affect the appearance of the picture to be stored.
///
/// Example:
/// ```
/// <p:nvPicPr>
///     <p:cNvPr id="4" name="lake.JPG" descr="Picture of a Lake" />
///     <p:cNvPicPr>
///         <a:picLocks noChangeAspect="1"/>
///     </p:cNvPicPr>
/// </p:nvPicPr>
/// ```
///
/// nvPicPr (Non-Visual Properties for a Picture)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNonVisualPictureProperties {
    // Child Elements
    // cNvPicPr (Non-Visual Picture Drawing Properties)
    pub non_visual_picture_drawing_properties: Option<XlsxNonVisualPictureDrawingProperties>,

    // cNvPr (Non-Visual Drawing Properties)
    pub non_visual_drawing_properties: Option<XlsxNonVisualDrawingProperties>,
}

impl XlsxNonVisualPictureProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            non_visual_picture_drawing_properties: None,
            non_visual_drawing_properties: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cNvPicPr" => {
                    properties.non_visual_picture_drawing_properties =
                        Some(XlsxNonVisualPictureDrawingProperties::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cNvPr" => {
                    properties.non_visual_drawing_properties =
                        Some(XlsxNonVisualDrawingProperties::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"nvPicPr" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at NonVisualPictureProperties: `nvPicPr`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
