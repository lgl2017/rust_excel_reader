use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::excel::XmlReader;

use crate::{helper::string_to_int, raw::drawing::color::ColorEnum};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.glow?view=openxml-3.0.1
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

#[derive(Debug, Clone, PartialEq)]
pub struct Glow {
    // children
    pub color: Option<ColorEnum>,

    // attribute
    pub rad: Option<i64>,
}

impl Glow {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
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
                            glow.rad = string_to_int(&string_value);
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

        glow.color = ColorEnum::load(reader, b"glow")?;

        Ok(glow)
    }
}
