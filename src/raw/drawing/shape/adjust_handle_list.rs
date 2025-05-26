use super::{adjust_handle_polar::XlsxAdjustHandlePolar, adjust_handle_xy::XlsxAdjustHandleXY};
use crate::excel::XmlReader;
use anyhow::bail;
use quick_xml::events::Event;
use std::io::Read;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.adjusthandlelist?view=openxml-3.0.1
///
/// This element specifies the adjust handles that are applied to a custom geometry.
///
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
pub type XlsxAdjustHandleList = Vec<XlsxAdjustHandleType>;

pub(crate) fn load_adjust_handle_list(
    reader: &mut XmlReader<impl Read>,
) -> anyhow::Result<XlsxAdjustHandleList> {
    let mut handles: Vec<XlsxAdjustHandleType> = vec![];

    let mut buf = Vec::new();

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ahPolar" => {
                handles.push(XlsxAdjustHandleType::Polar(XlsxAdjustHandlePolar::load(
                    reader, e,
                )?));
            }
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ahXY" => {
                handles.push(XlsxAdjustHandleType::XY(XlsxAdjustHandleXY::load(
                    reader, e,
                )?));
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"ahLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(handles)
}

#[derive(Debug, Clone, PartialEq)]
pub enum XlsxAdjustHandleType {
    Polar(XlsxAdjustHandlePolar),
    XY(XlsxAdjustHandleXY),
}
