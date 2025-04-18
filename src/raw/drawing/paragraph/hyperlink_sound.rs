use anyhow::bail;
use quick_xml::events::BytesStart;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.hyperlinksound?view=openxml-3.0.1
/// This element specifies a sound to be played when a hyperlink within the document is activated.
/// This sound is specified from within the parent hyperlink element.
///
/// Example:
/// ```
/// <snd embed="rId1" name="mysound/>
/// ```
// tag: snd
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxHyperlinkSound {
    // Attributes	Description
    // embed (Embedded Audio File Relationship ID)
    pub embed: Option<String>,

    /// Specifies the original name or given short name for the corresponding sound. This is used to distinguish this sound from others by providing a human readable name for the attached sound should the user need to identify the sound among others within the UI.
    // name (Sound Name)
    pub name: Option<String>,
}

impl XlsxHyperlinkSound {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut sound = Self {
            embed: None,
            name: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"embed" => {
                            sound.embed = Some(string_value);
                        }
                        b"name" => {
                            sound.name = Some(string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(sound)
    }
}
