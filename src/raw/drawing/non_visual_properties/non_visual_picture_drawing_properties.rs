use std::io::Read;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{excel::XmlReader, helper::string_to_bool};

use super::picture_locks::{load_picture_locks, XlsxPictureLocks};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.nonvisualpicturedrawingproperties?view=openxml-3.0.1
///
/// Example
/// ```
/// <p:cNvPicPr>
///     <a:picLocks noChangeAspect="1"/>
/// </p:cNvPicPr>
/// ```
///
/// cNvPicPr (Non-Visual Picture Drawing Properties)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNonVisualPictureDrawingProperties {
    // Child Elements
    // extLst (Extension List) Not supported
    /// picLocks (Picture Locks)
    pub picture_locks: Option<XlsxPictureLocks>,

    // Attributes
    /// preferRelativeResize (Relative Resize Preferred)
    ///
    /// Specifies if the user interface should show the resizing of the picture based on the picture's current size or its original size.
    /// If this attribute is set to true, then scaling is relative to the original picture size as opposed to the current picture size.
    pub prefer_relative_resize: Option<bool>,
}

impl XlsxNonVisualPictureDrawingProperties {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut properties = Self {
            picture_locks: None,
            prefer_relative_resize: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"preferRelativeResize" => {
                            properties.prefer_relative_resize = string_to_bool(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"picLocks" => {
                    properties.picture_locks = Some(load_picture_locks(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cNvPicPr" => break,
                Ok(Event::Eof) => {
                    bail!(
                        "unexpected end of file at NonVisualPictureDrawingProperties: `cNvPicPr`."
                    )
                }
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
