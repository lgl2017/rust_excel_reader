use std::io::Read;
use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::raw::drawing::font::{XlsxMajorFont, XlsxMinorFont};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fontscheme?view=openxml-3.0.1
/// This element defines the font scheme within the theme.
/// The font scheme consists of a pair of major and minor fonts for which to use in a document.
/// The major font corresponds well with the heading areas of a document, and the minor font corresponds well with the normal text or paragraph areas.
/// Example:
/// ```
/// <fontScheme name="sample">
///   <majorFont>
/// …  </majorFont>
///   <minorFont>
/// …  </minorFont>
/// </fontScheme>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxFontScheme {
    // child: extLst (Extension List)	Not supported

    /* Children */
    // majorFont (Major Font)	§20.1.4.1.24
    pub major_font: Option<XlsxMajorFont>,

    // minorFont (Minor fonts)
    pub minor_font: Option<XlsxMinorFont>,

    /* Attributes */
    // name (Name)
    pub name: Option<String>,
}

impl XlsxFontScheme {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut scheme = Self {
            major_font: None,
            minor_font: None,
            name: None,
        };

        let attributes = e.attributes();
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"name" => {
                            scheme.name = Some(string_value);
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

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"majorFont" => {
                    scheme.major_font = Some(XlsxMajorFont::load_major(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"minorFont" => {
                    scheme.minor_font = Some(XlsxMinorFont::load_minor(reader)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fontScheme" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(scheme)
    }
}
