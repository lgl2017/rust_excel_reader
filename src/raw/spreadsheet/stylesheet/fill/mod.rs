use anyhow::bail;
use gradient_fill::GradientFill;
use pattern_fill::PatternFill;
use quick_xml::events::Event;

use crate::excel::XmlReader;

pub mod gradient_fill;
pub mod pattern_fill;

/// Example
/// ```
/// <fills count="3">
///     <fill>
///         <patternFill patternType="none" />
///     </fill>
///     <fill>
///         <patternFill patternType="gray125" />
///     </fill>
///      </fill>
/// </fills>
/// ```
pub type Fills = Vec<Fill>;

pub(crate) fn load_fills(reader: &mut XmlReader) -> anyhow::Result<Fills> {
    let mut buf = Vec::new();
    let mut fills: Vec<Fill> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                let Some(fill) = Fill::load(reader)? else {
                    continue;
                };
                fills.push(fill);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fills" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(fills)
}

/// fill (Fill)
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fill?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub enum Fill {
    // children
    // xml tag name: patternFill
    PatternFill(PatternFill),
    // xml tag name: gradientFill
    GradientFill(GradientFill),
}

impl Fill {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Option<Self>> {
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"patternFill" => {
                    let pattern_fill = PatternFill::load(reader, e)?;
                    return Ok(Some(Fill::PatternFill(pattern_fill)));
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gradientFill" => {
                    let gradient_fill = GradientFill::load(reader, e)?;
                    return Ok(Some(Fill::GradientFill(gradient_fill)));
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fill" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(None)
    }
}
