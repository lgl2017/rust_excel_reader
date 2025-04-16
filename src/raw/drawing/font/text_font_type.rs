use anyhow::bail;
use quick_xml::events::BytesStart;

#[derive(Debug, Clone, PartialEq)]
pub struct TextFontType {
    // attributes
    /// Similar Character Set
    // charset
    pub charset: Option<String>,

    /// Panose Setting
    // panose
    pub panose: Option<String>,

    /// Similar Font Family
    // pitchFamily
    pub pitch_family: Option<String>,

    /// Text Typeface
    // typeface
    pub typeface: Option<String>,
}

impl TextFontType {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();

        let mut font = Self {
            charset: None,
            panose: None,
            pitch_family: None,
            typeface: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"charset" => {
                            font.charset = Some(string_value);
                        }
                        b"panose" => {
                            font.panose = Some(string_value);
                        }
                        b"pitchFamily" => {
                            font.pitch_family = Some(string_value);
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
