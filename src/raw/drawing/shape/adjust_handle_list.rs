use super::{adjust_handle_polar::XlsxAdjustHandlePolar, adjust_handle_xy::XlsxAdjustHandleXY};
use crate::excel::XmlReader;
use std::io::Read;
use anyhow::bail;
use quick_xml::events::Event;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.adjusthandlelist?view=openxml-3.0.1
///
/// This element specifies the adjust handles that are applied to a custom geometry.
/// These adjust handles specify points within the geometric shape that can be used to perform certain transform operations on the shape.
///
/// Example
/// ```
/// <a:ahLst>
///     <a:ahPolar gdRefAng="" gdRefR="">
///        <a:pos x="2" y="2"/>
///     </a:ahPolar>
///     <a:ahXY gdRefAng="" gdRefR="">
///        <a:pos x="2" y="2"/>
///     </a:ahXY>
/// </a:ahLst>
/// ```
// tag: ahLst
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxAdjustHandleList {
    // children
    pub adjust_handle_polar: Option<Vec<XlsxAdjustHandlePolar>>,
    pub adjust_handle_xy: Option<Vec<XlsxAdjustHandleXY>>,
}

impl XlsxAdjustHandleList {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut polars: Vec<XlsxAdjustHandlePolar> = vec![];
        let mut xys: Vec<XlsxAdjustHandleXY> = vec![];

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ahPolar" => {
                    polars.push(XlsxAdjustHandlePolar::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ahXY" => {
                    xys.push(XlsxAdjustHandleXY::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"ahLst" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(Self {
            adjust_handle_polar: Some(polars),
            adjust_handle_xy: Some(xys),
        })
    }
}
