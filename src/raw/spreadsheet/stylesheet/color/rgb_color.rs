use anyhow::bail;
use quick_xml::events::BytesStart;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.rgbcolor?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxRgbColor {
    pub rgb: Option<String>,
}

impl XlsxRgbColor {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut color = Self { rgb: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"rgb" => {
                            color.rgb = Some(string_value);
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
        Ok(color)
    }
}
