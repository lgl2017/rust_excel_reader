use anyhow::bail;
use quick_xml::events::BytesStart;
use std::io::Read;

use crate::excel::XmlReader;

use crate::helper::string_to_unsignedint;
use crate::raw::drawing::color::XlsxColorEnum;

/// innerShdw (Inner Shadow Effect
///
/// specifies an inner shadow effect.
/// A shadow is applied within the edges of the object according to the parameters given by the attributes
///
///  Example:
/// ```
/// <a:innerShdw blurRad="10" dir"90" dist="10">
///     <a:schemeClr val="phClr">
///         <a:lumMod val="99000" />
///         <a:satMod val="120000" />
///          a:shade val="78000" />
///     </a:schemeClr>
/// </a:innerShdw>
/// ```
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.innershadow?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxInnerShadow {
    // children
    pub color: Option<XlsxColorEnum>,

    // attribute
    /// blurRad
    ///
    /// Specifies the blur radius in emu.
    pub blur_rad: Option<u64>,

    /// Specifies the direction to offset the shadow as angle
    pub dir: Option<u64>,

    /// Specifies how far to offset the shadow
    pub dist: Option<u64>,
}

impl XlsxInnerShadow {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut shadow = Self {
            color: None,
            blur_rad: None,
            dir: None,
            dist: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"blurRad" => {
                            shadow.blur_rad = string_to_unsignedint(&string_value);
                        }
                        b"dir" => {
                            shadow.dir = string_to_unsignedint(&string_value);
                        }
                        b"dist" => {
                            shadow.dist = string_to_unsignedint(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        shadow.color = XlsxColorEnum::load(reader, b"innerShdw")?;

        Ok(shadow)
    }
}
