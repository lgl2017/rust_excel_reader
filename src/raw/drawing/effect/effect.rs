use anyhow::bail;
use quick_xml::events::BytesStart;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effect?view=openxml-3.0.1
///
/// specifies a reference to an existing effect container
// tag: effect
#[derive(Debug, Clone, PartialEq)]
pub struct Effect {
    // attributes
    /// Specifies the reference.
    ///
    /// Its value can be the name of an effect container, or one of four special references:
    /// - fill - refers to the fill effect
    /// - line - refers to the line effect
    /// - fillLine - refers to the combined fill and line effects
    /// - children - refers to the combined effects from logical child shapes or text
    pub r#ref: Option<String>,
}

impl Effect {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut effect = Self { r#ref: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"ref" => {
                            effect.r#ref = Some(string_value);
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

        Ok(effect)
    }
}
