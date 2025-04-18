use super::path_point::XlsxPoint;
use crate::excel::XmlReader;
use anyhow::bail;
use quick_xml::events::Event;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.moveto?view=openxml-3.0.1
///
/// This element specifies a set of new coordinates to move the shape cursor to.
///
/// Example
/// ```
/// <a:moveTo>
///     <a:pt x="0" y="1261641"/>
/// </a:moveTo>
/// ```
// tag: moveTo
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxMoveTo {
    // Child
    point: Option<XlsxPoint>,
}

impl XlsxMoveTo {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut point: Option<XlsxPoint> = None;

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pt" => {
                    point = Some(XlsxPoint::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"moveTo" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(Self { point })
    }
}
