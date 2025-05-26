use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{helper::string_to_unsignedint, raw::drawing::st_types::STPositiveCoordinate};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bevelbottom?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:sp3d>
///     <a:bevelB w="139700" h="127000" prst="coolSlant"/>
/// </a:sp3d>
/// ```
pub type XlsxBevelBottom = XlsxBevel;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.beveltop?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:sp3d>
///     <a:bevelT w="139700" h="127000" prst="coolSlant"/>
/// </a:sp3d>
/// ```
pub type XlsxBevelTop = XlsxBevel;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bevel?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxBevel {
    // Attributes
    /// Specifies the height of the bevel, or how far above the shape it is applied.
    // h (Height)
    pub h: Option<STPositiveCoordinate>,

    /// Specifies the width of the bevel, or how far into the shape it is applied.
    // w (Width)
    pub w: Option<STPositiveCoordinate>,

    /// Specifies the preset bevel type which defines the look of the bevel.
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bevelpresetvalues?view=openxml-3.0.1
    // prst (Preset Bevel)
    pub prst: Option<String>,
}

impl XlsxBevel {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut bevel = Self {
            h: None,
            w: None,
            prst: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"h" => bevel.h = string_to_unsignedint(&string_value),
                        b"w" => bevel.w = string_to_unsignedint(&string_value),
                        b"prst" => bevel.prst = Some(string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(bevel)
    }
}
