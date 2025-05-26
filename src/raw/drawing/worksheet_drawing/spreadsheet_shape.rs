use std::io::Read;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    helper::string_to_bool,
    raw::drawing::{
        non_visual_properties::non_visual_shape_properties::XlsxNonVisualShapeProperties,
        shape::{shape_properties::XlsxShapeProperties, shape_style::XlsxShapeStyle},
        text::shape_text_body::XlsxShapeTextBody,
    },
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.shape?view=openxml-3.0.1
///
/// NOTE: This class is different from a [Drawing.Shape](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shape?view=openxml-3.0.1).
///
/// This element specifies the existence of a single shape.
/// A shape can either be a preset or a custom geometry, defined using the DrawingML framework.
/// In addition to a geometry each shape can have both visual and non-visual properties attached.
/// Text and corresponding styling information can also be attached to a shape.
/// This shape is specified along with all other shapes within either the shape tree or group shape elements.
///
/// Example:
/// ```
/// <p:sp macro="" textlink="$D$3">
///   <p:nvSpPr>
///     <p:cNvPr id="2" name="Rectangle 1"/>
///     <p:cNvSpPr>
///       <a:spLocks noGrp="1"/>
///     </p:cNvSpPr>
///   </p:nvSpPr>
/// …</p:sp>
/// ```
///
/// sp (Shape)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxShape {
    // Child Elements
    // extLst (Extension List)

    // nvSpPr (Non-Visual Properties for a Shape)
    pub non_visual_shape_properties: Option<XlsxNonVisualShapeProperties>,

    // spPr (Shape Properties)	§20.1.2.2.35
    pub shape_properties: Option<XlsxShapeProperties>,

    // style (Shape Style)	§20.1.2.2.37
    pub shape_style: Option<XlsxShapeStyle>,

    // txBody (Shape Text Body)
    pub text_body: Option<XlsxShapeTextBody>,

    // attributes
    /// LockText
    ///
    /// Lock Text Flag
    ///
    /// default to true
    pub lock_text: Option<bool>,

    /// Macro
    ///
    /// Reference to Custom Function
    pub r#macro: Option<String>,

    /// Published
    ///
    /// Publish to Server Flag
    pub published: Option<bool>,

    /// TextLink
    ///
    /// Reference used by a fld (TextField) of type `TxLink`.
    ///
    /// Example:
    /// ```
    /// textlink="$D$3"
    /// ```
    pub text_link: Option<String>,
}

impl XlsxShape {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut shape = Self {
            non_visual_shape_properties: None,
            shape_properties: None,
            shape_style: None,
            text_body: None,
            lock_text: None,
            r#macro: None,
            published: None,
            text_link: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"textlink" => {
                            shape.text_link = Some(string_value);
                        }
                        b"fLocksText" => {
                            shape.lock_text = string_to_bool(&string_value);
                        }
                        b"macro" => {
                            shape.r#macro = Some(string_value);
                        }
                        b"fPublished" => {
                            shape.published = string_to_bool(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"nvSpPr" => {
                    shape.non_visual_shape_properties =
                        Some(XlsxNonVisualShapeProperties::load(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spPr" => {
                    shape.shape_properties = Some(XlsxShapeProperties::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"style" => {
                    shape.shape_style = Some(XlsxShapeStyle::load(reader)?);
                }
                // sometimes directly as txBody
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"txBody" => {
                    shape.text_body = Some(XlsxShapeTextBody::load(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sp" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `sp`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        return Ok(shape);
    }
}
