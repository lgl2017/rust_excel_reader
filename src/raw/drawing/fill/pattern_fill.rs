use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::excel::XmlReader;

use crate::raw::drawing::color::XlsxColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.backgroundcolor?view=openxml-3.0.1
pub type XlsxBackgroundColor = XlsxColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.foregroundcolor?view=openxml-3.0.1
pub type XlsxForegroundColor = XlsxColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.patternfill?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:pattFill prst="cross">
///     <a:bgClr>
///         <a:solidFill>
///             <a:schemeClr val="phClr" />
///         </a:solidFill>
///     </a:bgClr>
///     <a:fgClr>
///         <a:schemeClr val="phClr">
///             <a:satMod val="110000" />
///             <a:lumMod val="100000" />
///             <a:shade val="100000" />
///         </a:schemeClr>
///     </a:fgClr>
/// </a:pattFill>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxPatternFill {
    // Child Elements
    /// bgClr (Background color)
    pub bg_clr: Option<XlsxBackgroundColor>,

    /// fgClr (Foreground color)
    pub fg_clr: Option<XlsxForegroundColor>,

    // Attributes
    /// Specifies one of a set of preset patterns to fill the object
    /// Allowed values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetpatternvalues?view=openxml-3.0.1
    // tag: prst (Preset Pattern)
    pub prst: Option<String>,
}

impl XlsxPatternFill {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut fill = Self {
            bg_clr: None,
            fg_clr: None,
            prst: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"prst" => {
                            fill.prst = Some(string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"bgClr" => {
                    fill.bg_clr = XlsxBackgroundColor::load(reader, b"bgClr")?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fgClr" => {
                    fill.fg_clr = XlsxForegroundColor::load(reader, b"fgClr")?;
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"pattFill" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(fill)
    }
}
