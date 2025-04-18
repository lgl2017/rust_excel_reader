use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::excel::XmlReader;
use crate::helper::string_to_unsignedint;
use crate::raw::drawing::color::XlsxColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linereference?view=openxml-3.0.1
///
/// This element defines a reference to a line style within the style matrix.
/// The idx attribute refers the index of a line style within the `fillStyleLst` element.
///
/// Example
/// ```
/// <lnRef idx="1">
///     <schemeClr val="accent2"/>
/// </lnRef>
/// ```
// tag: lnRef
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxLineReference {
    // Child Elements
    color: Option<XlsxColorEnum>,

    // Attributes	Description
    /// Specifies the style matrix index of the style referred to
    // tag: idx
    pub index: Option<u64>,
}

impl XlsxLineReference {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut reference = Self {
            color: XlsxColorEnum::load(reader, b"lnRef")?,
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
