use anyhow::bail;
use quick_xml::events::BytesStart;
use std::io::Read;

use crate::excel::XmlReader;
use crate::helper::string_to_unsignedint;
use crate::raw::drawing::color::XlsxColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillreference?view=openxml-3.0.1
///
/// This element defines a reference to a fill style within the style matrix.
///
/// idx:
/// - value of 0 or 1000 indicates no background,
/// - values 1-999 refer to the index of a fill style within the fillStyleLst element,
/// - values 1001 and above refer to the index of a background fill style within the bgFillStyleLst element.
///     For example: The value 1001 corresponds to the first background fill style, 1002 to the second background fill style, and so on.
///
/// Example
/// ```
/// <fillRef idx="0">
///     <schemeClr val="accent2"/>
/// </fillRef>
/// ```
// tag: fillRef
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxFillReference {
    // Child Elements
    pub color: Option<XlsxColorEnum>,

    // Attributes	Description
    /// Specifies the style matrix index of the style referred to
    // tag: idx
    pub index: Option<u64>,
}

impl XlsxFillReference {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut reference = Self {
            color: XlsxColorEnum::load(reader, b"fillRef")?,
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
