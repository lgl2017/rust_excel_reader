use std::io::Read;

use anyhow::bail;
use quick_xml::events::Event;

use crate::excel::XmlReader;

use super::{
    non_visual_drawing_properties::XlsxNonVisualDrawingProperties,
    non_visual_ink_content_part_properties::XlsxNonVisualInkContentPartProperties,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.excel.drawing.excelnonvisualcontentpartshapeproperties?view=openxml-3.0.1
///
/// A complex type that specifies non-visual properties of a contentPart element
///
/// xdr14:nvPr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxExcelNonVisualContentPartShapeProperties {
    // Child Elements:
    /// cNvPr:
    ///
    /// A CT_NonVisualDrawingProps element ([ISO/IEC-29500-1] section A.4.1) that specifies the non-visual drawing properties of the content part.
    /// This enables additional information that does not affect the appearance of the content part to be stored.
    pub non_visual_drawing_properties: Option<XlsxNonVisualDrawingProperties>,

    /// cNvContentPartPr:
    ///
    /// A CT_NonVisualInkContentPartProperties element that specifies non-visual ink properties of the content part.
    /// This enables additional information that does not affect the appearance of ink in the content part to be stored.
    pub non_visual_ink_content_part_properties: Option<XlsxNonVisualInkContentPartProperties>,
}

impl XlsxExcelNonVisualContentPartShapeProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            non_visual_drawing_properties: None,
            non_visual_ink_content_part_properties: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cNvPr" => {
                    properties.non_visual_drawing_properties =
                        Some(XlsxNonVisualDrawingProperties::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cNvContentPartPr" => {
                    properties.non_visual_ink_content_part_properties =
                        Some(XlsxNonVisualInkContentPartProperties::load(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"nvPr" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at XlsxExcelNonVisualContentPartShapeProperties: `nvPr`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
