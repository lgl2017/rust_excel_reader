use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::helper::string_to_int;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.customdash?view=openxml-3.0.1
///
/// This element specifies a custom dashing scheme.
/// It is a list of dash stop elements which represent building block atoms upon which the custom dashing scheme is built.

#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCustomDash {
    // children
    /// A list of dash stops
    pub ds: Option<Vec<XlsxDashStop>>,
}

impl XlsxCustomDash {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut stops: Vec<XlsxDashStop> = vec![];

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ds" => {
                    stops.push(XlsxDashStop::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"custDash" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(Self { ds: Some(stops) })
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.dashstop?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxDashStop {
    // Attributes
    /// Specifies the length of the dash relative to the line width.
    // d (Dash Length)
    pub d: Option<i64>,

    /// Specifies the length of the space relative to the line width.
    // sp (Space Length)
    pub sp: Option<i64>,
}

impl XlsxDashStop {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();

        let mut dash_stop = Self { d: None, sp: None };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"d" => dash_stop.d = string_to_int(&string_value),
                        b"sp" => dash_stop.sp = string_to_int(&string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(dash_stop)
    }
}
