use std::io::Read;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{excel::XmlReader, helper::string_to_bool};

use super::shape_locks::XlsxShapeLocks;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.nonvisualshapedrawingproperties?view=openxml-3.0.1
///
/// This element specifies the non-visual drawing properties for a shape.
/// These properties are to be used by the generating application to determine how the shape should be dealt with
///
/// Example:
/// ```
/// <p:sp>
///   <p:nvSpPr>
///     <p:cNvPr id="2" name="Rectangle 1"/>
///     <p:cNvSpPr>
///       <a:spLocks noGrp="1"/>
///     </p:cNvSpPr>
///   </p:nvSpPr>
/// …</p:sp>
/// ```
/// cNvSpPr (Non-Visual Shape Drawing Properties)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNonVisualShapeDrawingProperties {
    // Child Elements
    // extLst (Extension List)	§20.1.2.2.15
    // spLocks (Shape Locks)
    pub shape_locks: Option<XlsxShapeLocks>,

    // Attributes
    /// txBox (Text Box)
    ///
    /// Specifies that the corresponding shape is a text box and thus should be treated as such by the generating application.
    /// If this attribute is omitted then it is assumed that the corresponding shape is not specifically a text box.
    pub text_box: Option<bool>,
}

impl XlsxNonVisualShapeDrawingProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            shape_locks: None,
            text_box: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"txBox" => {
                            properties.text_box = string_to_bool(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"spLocks" => {
                    properties.shape_locks = Some(XlsxShapeLocks::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cNvSpPr" => break,
                Ok(Event::Eof) => {
                    bail!("unexpected end of file at NonVisualPictureDrawingProperties: `cNvSpPr`.")
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
