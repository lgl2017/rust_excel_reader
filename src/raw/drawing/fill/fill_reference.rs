use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::excel::XmlReader;

use crate::raw::drawing::color::ColorEnum;
use crate::helper::string_to_unsignedint;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillreference?view=openxml-3.0.1
///
/// This element defines a reference to a fill style within the style matrix.
///
/// Example
/// ```
/// <fillRef idx="0">
///     <schemeClr val="accent2"/>
/// </fillRef>
/// ```
// tag: fillRef
#[derive(Debug, Clone, PartialEq)]
pub struct FillReference {
    // Child Elements
    color: Option<ColorEnum>,

    // Attributes	Description
    /// Specifies the style matrix index of the style referred to
    // tag: idx
    pub index: Option<u64>,
}

impl FillReference {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut reference = Self {
            color: ColorEnum::load(reader, b"fillRef")?,
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
