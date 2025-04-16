use anyhow::bail;
use quick_xml::events::BytesStart;

#[derive(Debug, Clone, PartialEq)]
pub struct BaseFont {
    pub typeface: Option<String>,
}

impl BaseFont {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut font = Self { typeface: None };
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"typeface" => {
                            font.typeface = Some(string_value);
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

        return Ok(font);
    }
}
