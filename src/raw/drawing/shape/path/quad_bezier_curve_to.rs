use super::path_point::XlsxPoint;
use crate::excel::XmlReader;
use std::io::Read;
use anyhow::bail;
use quick_xml::events::Event;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.quadraticbeziercurveto?view=openxml-3.0.1
///
/// This element specifies to draw a quadratic bezier curve along the specified points.
///
/// To specify a quadratic bezier curve there needs to be 2 points specified.
/// The first is a control point used in the quadratic bezier calculation and the last is the ending point for the curve.
// tag: quadBezTo
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxQuadraticBezierCurveTo {
    // Child
    points: Option<Vec<XlsxPoint>>,
}

impl XlsxQuadraticBezierCurveTo {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut points: Vec<XlsxPoint> = vec![];

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"pt" => {
                    points.push(XlsxPoint::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"quadBezTo" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }
        Ok(Self {
            points: Some(points),
        })
    }
}
