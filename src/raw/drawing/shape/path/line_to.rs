use super::path_point::XlsxPoint;
use crate::excel::XmlReader;
use anyhow::bail;
use quick_xml::events::Event;
use std::io::Read;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.lineto?view=openxml-3.0.1
///
/// This element specifies the drawing of a straight line from the current pen position to the new point specified.
///
/// Example
/// ```
/// <a:lnTo>
///     <a:pt x="2650602" y="1261641"/>
/// </a:lnTo>
/// ```
// tag: lnTo
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxLineTo {
    // Child
    pub point: Option<XlsxPoint>,
}

impl XlsxLineTo {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut point: Option<XlsxPoint> = None;

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pt" => {
                    point = Some(XlsxPoint::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"lnTo" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(Self { point })
    }
}
