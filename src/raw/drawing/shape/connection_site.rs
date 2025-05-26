use anyhow::bail;
use quick_xml::events::{BytesStart, Event};
use std::io::Read;

use crate::{excel::XmlReader, raw::drawing::st_types::STAdjustAngle};

use super::position::XlsxPosition;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.connectionsite?view=openxml-3.0.1
///
/// This element specifies the existence of a connection site on a custom shape.
///
/// Example
/// ```
/// <a:cxnLst>
///     <a:cxn ang="0">
///         <a:pos x="0" y="679622"/>
///     </a:cxn>
///     <a:cxn ang="0">
///         <a:pos x="1705233" y="679622"/>
///     </a:cxn>
/// </a:cxnLst>
/// ```
// tag: cxn
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxConnectionSite {
    // Child Elements
    // pos (Shape Position Coordinate)
    pub position: Option<XlsxPosition>,

    // Attributes
    /// Specifies the incoming connector angle.
    ///  This angle is the angle around the connection site that an incoming connector tries to be routed to.
    ///  This allows connectors to know where the shape is in relation to the connection site and route connectors so as to avoid any overlap with the shape.
    // ang (Connection Site Angle)
    pub angle: Option<STAdjustAngle>,
}

impl XlsxConnectionSite {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut site = Self {
            position: None,
            angle: None,
        };

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"ang" => {
                            site.angle = Some(STAdjustAngle::from_string(&string_value));
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pos" => {
                    site.position = Some(XlsxPosition::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cxn" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(site)
    }
}
