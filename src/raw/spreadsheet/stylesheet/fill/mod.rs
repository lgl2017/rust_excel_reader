use anyhow::bail;
use gradient_fill::XlsxGradientFill;
use pattern_fill::XlsxPatternFill;
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
pub type XlsxFills = Vec<XlsxFill>;

pub(crate) fn load_fills(reader: &mut XmlReader) -> anyhow::Result<XlsxFills> {
    let mut buf = Vec::new();
    let mut fills: Vec<XlsxFill> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fill" => {
                let Some(fill) = XlsxFill::load(reader)? else {
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
pub enum XlsxFill {
    // children
    // xml tag name: patternFill
    PatternFill(XlsxPatternFill),
    // xml tag name: gradientFill
    GradientFill(XlsxGradientFill),
}

impl XlsxFill {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Option<Self>> {
        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"patternFill" => {
                    let pattern_fill = XlsxPatternFill::load(reader, e)?;
                    return Ok(Some(XlsxFill::PatternFill(pattern_fill)));
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gradientFill" => {
                    let gradient_fill = XlsxGradientFill::load(reader, e)?;
                    return Ok(Some(XlsxFill::GradientFill(gradient_fill)));
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
