use std::io::Read;
use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::excel::XmlReader;

use crate::{helper::string_to_int, raw::drawing::color::XlsxColorEnum};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetshadow?view=openxml-3.0.1
/// specifies that a preset shadow is to be used.
/// Each preset shadow is equivalent to a specific outer shadow effect.
/// For each preset shadow, the color element, direction attribute, and distance attribute represent the color, direction, and distance parameters of the corresponding outer shadow.
/// Additionally, the rotateWithShape attribute of corresponding outer shadow is always false. Other non-default parameters of the outer shadow are dependent on the prst attribute
///
///  Example:
/// ```
/// <a:prstShdw dir"90" dist="10" prst="shdw19">
///     <a:schemeClr val="phClr">
///         <a:lumMod val="99000" />
///         <a:satMod val="120000" />
///          a:shade val="78000" />
///     </a:schemeClr>
/// </a:prstShdw>
/// ```

#[derive(Debug, Clone, PartialEq)]
pub struct XlsxPresetShadow {
    // children
    pub color: Option<XlsxColorEnum>,

    // attribute
    /// Specifies the direction to offset the shadow as angle
    pub dir: Option<i64>,

    /// Specifies how far to offset the shadow
    pub dist: Option<i64>,

    /// Specifies which preset shadow to use
    /// allowed values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetshadowvalues?view=openxml-3.0.1
    pub prst: Option<String>,
}

impl XlsxPresetShadow {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut shadow = Self {
            color: None,
            prst: None,
            dir: None,
            dist: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"prst" => shadow.prst = Some(string_value),
                        b"dir" => {
                            shadow.dir = string_to_int(&string_value);
                        }
                        b"dist" => {
                            shadow.dist = string_to_int(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        shadow.color = XlsxColorEnum::load(reader, b"prstShdw")?;

        Ok(shadow)
    }
}
