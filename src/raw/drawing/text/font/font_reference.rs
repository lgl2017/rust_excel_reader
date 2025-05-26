use anyhow::bail;
use quick_xml::events::BytesStart;
use std::io::Read;

use crate::excel::XmlReader;

use crate::raw::drawing::color::XlsxColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fontreference?view=openxml-3.0.1
///
/// This element represents a reference to a themed font.
/// When used it specifies which themed font to use along with a choice of color.
///
/// Example
/// ```
/// <fontRef idx="minor">
///     <schemeClr val="tx1"/>
/// </fontRef>
/// ```
// tag: fontRef
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxFontReference {
    // Child Elements
    pub color: Option<XlsxColorEnum>,

    // Attributes
    /// Specifies the identifier of the font to reference.
    ///
    /// Allowed Values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fontcollectionindexvalues?view=openxml-3.0.1
    /// - `major`
    /// - `minor`
    /// - `none`
    // tag: idx
    pub index: Option<String>,
}

impl XlsxFontReference {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut reference = Self {
            color: XlsxColorEnum::load(reader, b"fontRef")?,
            index: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"idx" => {
                            reference.index = Some(string_value);
                            break;
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(reference)
    }
}
