use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::excel::XmlReader;

use super::{
    non_visual_connector_shape_drawing_properties::XlsxNonVisualConnectorShapeDrawingProperties,
    non_visual_drawing_properties::XlsxNonVisualDrawingProperties,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.nonvisualconnectionshapeproperties?view=openxml-3.0.1
///
///
/// This element specifies all non-visual properties for a connection shape.
/// This element is a container for the non-visual identification properties, shape properties and application properties that are to be associated with a connection shape.
/// This allows for additional information that does not affect the appearance of the connection shape to be stored.
///
/// Example
/// ```
/// <p:cxnSp>
/// <p:nvCxnSpPr>
///   <p:cNvPr id="3" name="Elbow Connector 3"/>
///   <p:cNvCxnSpPr>
///     <a:stCxn id="1" idx="3"/>
///     <a:endCxn id="2" idx="1"/>
///   </p:cNvCxnSpPr>
///   <p:nvPr/>
/// </p:nvCxnSpPr>
/// …  </p:cxnSp>
/// ```
///
/// nvCxnSpPr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNonVisualConnectionShapeProperties {
    // Child Elements
    // cNvCxnSpPr (Non-Visual Connector Shape Drawing Properties)
    pub non_visual_connector_shape_drawing_properties:
        Option<XlsxNonVisualConnectorShapeDrawingProperties>,

    // cNvPr (Non-Visual Drawing Properties)
    pub non_visual_drawing_properties: Option<XlsxNonVisualDrawingProperties>,
}

impl XlsxNonVisualConnectionShapeProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            non_visual_connector_shape_drawing_properties: None,
            non_visual_drawing_properties: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cNvCxnSpPr" => {
                    properties.non_visual_connector_shape_drawing_properties =
                        Some(XlsxNonVisualConnectorShapeDrawingProperties::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cNvPr" => {
                    properties.non_visual_drawing_properties =
                        Some(XlsxNonVisualDrawingProperties::load(reader, e)?);
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"nvCxnSpPr" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at XlsxNonVisualConnectionShapeProperties: `nvCxnSpPr`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
