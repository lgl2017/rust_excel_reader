use anyhow::bail;
use quick_xml::events::BytesStart;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.supplementalfont?view=openxml-3.0.1
///
/// Example:
/// ```
/// <font script="Thai" typeface="Cordia New"/>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct SupplementalFont {
    // Attributes
    /// Specifies the script, or language, in which the typeface is supposed to be used.
    // script (Script)
    pub script: Option<String>,

    /// Specifies the font face to use.
    // typeface (Typeface)
    pub typeface: Option<String>,
}

impl SupplementalFont {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();

        let mut font = Self {
            script: None,
            typeface: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"script" => {
                            font.script = Some(string_value);
                        }
                        b"typeface" => {
                            font.typeface = Some(string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(font)
    }
}
