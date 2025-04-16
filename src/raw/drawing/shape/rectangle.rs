use anyhow::bail;
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rectangle?view=openxml-3.0.1
use quick_xml::events::BytesStart;

use crate::helper::string_to_int;

#[derive(Debug, Clone, PartialEq)]
pub struct Rectangle {
    // attributes
    /// Specifies the bottom edge of the rectangle.
    pub b: Option<i64>,

    /// Specifies the left edge of the rectangle
    pub l: Option<i64>,

    /// Specifies the right edge of the rectangle
    pub r: Option<i64>,

    /// Specifies the top edge of the rectangle
    pub t: Option<i64>,
}

impl Rectangle {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut rect = Self {
            b: None,
            l: None,
            r: None,
            t: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"b" => {
                            rect.b = string_to_int(&string_value);
                        }
                        b"l" => {
                            rect.l = string_to_int(&string_value);
                        }
                        b"r" => {
                            rect.r = string_to_int(&string_value);
                        }
                        b"t" => {
                            rect.t = string_to_int(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(rect)
    }
}
