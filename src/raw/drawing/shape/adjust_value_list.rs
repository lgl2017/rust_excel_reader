use super::shape_guide::XlsxShapeGuide;
use crate::excel::XmlReader;
use anyhow::bail;
use quick_xml::events::Event;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.adjustvaluelist?view=openxml-3.0.1
/// This element specifies the adjust values that are applied to the specified shape.
/// An adjust value is simply a guide that has a value based formula specified.
/// Example:
/// ```
/// <a:avLst>
///     <a:gd name="myGuide" fmla="val 2"/>
/// </a:avLst>
/// ```
// tag: avLst: List of Shape Adjust Values
pub type XlsxAdjustValueList = Vec<XlsxShapeGuide>;

pub(crate) fn load_adjust_value_list(
    reader: &mut XmlReader,
) -> anyhow::Result<XlsxAdjustValueList> {
    let mut buf = Vec::new();
    let mut guides: Vec<XlsxShapeGuide> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gd" => {
                guides.push(XlsxShapeGuide::load(e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"avLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    Ok(guides)
}
