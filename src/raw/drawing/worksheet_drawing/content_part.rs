use std::io::Read;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    raw::drawing::{
        non_visual_properties::{
            application_non_visual_drawing_properties::XlsxApplicationNonVisualDrawingProperties,
            excel_non_visual_content_part_shape_properties::XlsxExcelNonVisualContentPartShapeProperties,
        },
        shape::transform_2d::XlsxTransform2D,
    },
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.contentpart?view=openxml-3.0.1
///
/// This element specifies a reference to XML content in a format not defined by /IEC 29500, such as:
/// - MathML (http://www.w3.org/TR/MathML2/)
/// - SMIL (http://www.w3.org/TR/REC-smil/)
/// - SVG (http://www.w3.org/TR/SVG11/)
///
/// Example:
/// ```
/// <wsDr>
///   <twoCellAnchor>
///     <from>
///       <col>3</col>
///       <colOff>152400</colOff>
///       <row>5</row>
///       <rowOff>123825</rowOff>
///     </from>
///     <to>
///       <col>8</col>
///       <colOff>266700</colOff>
///       <row>22</row>
///       <rowOff>38100</rowOff>
///     </to>
///   </twoCellAnchor>
///   <contentPart r:id="svg1"/>
/// </wsDr>
/// ```
///
/// contentPart
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxContentPart {
    // children elements

    // extLst (Not Supported)
    /// nvPr (ApplicationNonVisualDrawingProperties)
    pub application_non_visual_drawing_properties:
        Option<XlsxApplicationNonVisualDrawingProperties>,

    /// nvContentPartPr (ExcelNonVisualContentPartShapeProperties)
    pub excel_non_visual_content_part_shape_properties:
        Option<XlsxExcelNonVisualContentPartShapeProperties>,

    /// xfrm (Transform2D)
    pub transform_2d: Option<XlsxTransform2D>,

    // attributes
    /// bwMode (BlackWhiteMode)
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blackwhitemodevalues?view=openxml-3.0.1
    pub black_white_mode: Option<String>,

    /// id (RelationshipId)
    pub id: Option<String>,
}

impl XlsxContentPart {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            application_non_visual_drawing_properties: None,
            excel_non_visual_content_part_shape_properties: None,
            transform_2d: None,
            black_white_mode: None,
            id: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"bwMode" => {
                            properties.black_white_mode = Some(string_value);
                        }
                        b"id" => {
                            properties.id = Some(string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"nvPr" => {
                    properties.application_non_visual_drawing_properties =
                        Some(XlsxApplicationNonVisualDrawingProperties::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"nvContentPartPr" => {
                    properties.excel_non_visual_content_part_shape_properties =
                        Some(XlsxExcelNonVisualContentPartShapeProperties::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"xfrm" => {
                    properties.transform_2d = Some(XlsxTransform2D::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"contentPart" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at XlsxContentPart: `contentPart`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
