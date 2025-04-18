use anyhow::bail;
use quick_xml::events::BytesStart;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.characterbullet?view=openxml-3.0.1
///
/// This element specifies that a character be applied to a set of bullets.
///
/// Example:
/// ```
/// <a:pPr …>
///     <a:buFont typeface="Calibri"/>
///     <a:buChar char="g"/>
/// </a:pPr>
/// ```
// tag: buChar
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCharacterBullet {
    // Attributes
    ///	Specifies the character to be used in place of the standard bullet point.
    /// This character can be any character for the specified font that is supported by the system upon which this document is being viewed.
    // char (Bullet Character)
    pub char: Option<String>,
}

impl XlsxCharacterBullet {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut bullet = Self { char: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"char" => {
                            bullet.char = Some(string_value);
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

        Ok(bullet)
    }
}
