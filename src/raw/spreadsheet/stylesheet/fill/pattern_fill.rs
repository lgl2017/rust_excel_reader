use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    raw::spreadsheet::stylesheet::color::{BackgroundColor, Color, ForegroundColor},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.patternfill?view=openxml-3.0.1
/// Example:
/// ```
/// <fill>
///     <patternFill patternType="solid">
///         <fgColor indexed="12" />
///         <bgColor auto="1" />
///     </patternFill>
///// </fill>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct PatternFill {
    // attributes
    /// patternType
    ///
    /// Specifies the fill pattern type (including solid and none).
    /// Default is none, when missing
    ///
    /// Allowed value: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.patternvalues?view=openxml-3.0.1.
    pub pattern_type: Option<String>,

    // children
    // xml tag name: fgColor
    /// foreground color
    pub foreground_color: Option<ForegroundColor>,

    // xml tag name: bgcolor
    /// background color
    pub background_color: Option<BackgroundColor>,
}

impl PatternFill {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut fill = Self {
            pattern_type: None,
            foreground_color: None,
            background_color: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"patternType" => {
                            fill.pattern_type = Some(string_value);
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

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fgColor" => {
                    let color: Color = Color::load(e)?;
                    fill.foreground_color = Some(color);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"bgColor" => {
                    let color: Color = Color::load(e)?;
                    fill.background_color = Some(color);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"patternFill" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(fill)
    }
}
