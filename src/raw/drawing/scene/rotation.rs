use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{helper::string_to_int, raw::drawing::st_types::STAngle};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rotation?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:rot lat="0" lon="0" rev="6000000"/>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxRotation {
    // attributes
    /// latitude coordinate (y rotation)
    pub lat: Option<STAngle>,

    /// longitude coordinate, (x rotation)
    pub long: Option<STAngle>,

    /// revolution about the axis as the latitude and longitude coordinates (z rotation)
    pub rev: Option<STAngle>,
}

impl XlsxRotation {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut rotation = Self {
            lat: None,
            long: None,
            rev: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"lat" => rotation.lat = string_to_int(&string_value),
                        b"long" => rotation.long = string_to_int(&string_value),
                        b"rev" => rotation.rev = string_to_int(&string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(rotation)
    }
}
