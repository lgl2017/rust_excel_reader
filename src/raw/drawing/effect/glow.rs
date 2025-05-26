use anyhow::bail;
use quick_xml::events::BytesStart;
use std::io::Read;

use crate::excel::XmlReader;

use crate::helper::string_to_unsignedint;
use crate::raw::drawing::color::XlsxColorEnum;

/// glow (Glow Effect)
///
/// specifies a glow effect, in which a color blurred outline is added outside the edges of the object.
///
/// Example:
/// ```
/// <a:glow rad="10">
///     <a:schemeClr val="phClr">
///         <a:lumMod val="99000" />
///         <a:satMod val="120000" />
///          a:shade val="78000" />
///     </a:schemeClr>
/// </a:glow>
/// ```
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.glow?view=openxml-3.0.1

#[derive(Debug, Clone, PartialEq)]
pub struct XlsxGlow {
    // children
    pub color: Option<XlsxColorEnum>,

    // attribute
    /// radius
    ///
    /// Specifies the radius of the glow in emu
    pub rad: Option<u64>,
}

impl XlsxGlow {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut glow = Self {
            color: None,
            rad: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"pos" => {
                            glow.rad = string_to_unsignedint(&string_value);
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

        glow.color = XlsxColorEnum::load(reader, b"glow")?;

        Ok(glow)
    }
}
