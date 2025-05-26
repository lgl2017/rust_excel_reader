use anyhow::bail;
use quick_xml::events::BytesStart;
use std::io::Read;

use crate::excel::XmlReader;
use crate::helper::string_to_unsignedint;
use crate::raw::drawing::color::XlsxColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectreference?view=openxml-3.0.1
///
/// This element defines a reference to an effect style within the style matrix.
///
/// Example
/// ```
/// // idx: 0 based
/// <effectRef idx="0">
///     <schemeClr val="accent2"/>
/// </effectRef>
/// ```
// tag: effectRef
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxEffectReference {
    // Child Elements
    pub color: Option<XlsxColorEnum>,

    // Attributes
    /// Specifies the style matrix 0 based index of the style referred to
    // tag: idx
    pub index: Option<u64>,
}

impl XlsxEffectReference {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut reference = Self {
            color: XlsxColorEnum::load(reader, b"effectRef")?,
            index: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"idx" => {
                            reference.index = string_to_unsignedint(&string_value);
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
