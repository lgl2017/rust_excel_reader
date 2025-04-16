use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::{string_to_bool, string_to_int};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blur?view=openxml-3.0.1
/// a blur effect that is applied to the entire shape, including its fill.
/// All color channels, including alpha, are affected.
#[derive(Debug, Clone, PartialEq)]
pub struct Blur {
    // attributes
    /// Specifies whether the bounds of the object should be grown as a result of the blurring.
    /// True indicates the bounds are grown while false indicates that they are not.
    // tag: grow (Grow Bounds)
    pub grow: Option<bool>,

    /// Specifies the radius of blur.
    // tag: rad (Radius)
    pub rad: Option<i64>,
}

impl Blur {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut blur = Self {
            grow: None,
            rad: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"grow" => {
                            blur.grow = string_to_bool(&string_value);
                        }
                        b"rad" => {
                            blur.rad = string_to_int(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(blur)
    }
}
