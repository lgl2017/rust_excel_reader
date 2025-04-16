use anyhow::bail;
use quick_xml::events::Event;
use crate::excel::XmlReader;
use super::shape_guide::ShapeGuide;

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
pub type AdjustValueList = Vec<ShapeGuide>;

pub(crate) fn load_adjust_value_list(
    reader: &mut XmlReader,
) -> anyhow::Result<AdjustValueList> {
    let mut buf = Vec::new();
    let mut guides: Vec<ShapeGuide> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gd" => {
                guides.push(ShapeGuide::load(e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"avLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    Ok(guides)
}

// #[derive(Debug, Clone, PartialEq)]
// pub struct AdjustValueList {
//     // children
//     // gd (Shape Guide)
//     pub guide: Option<Vec<ShapeGuide>>,
// }

// impl AdjustValueList {
//     pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
//         let mut buf = Vec::new();
//         let mut guides: Vec<ShapeGuide> = vec![];

//         loop {
//             buf.clear();

//             match reader.read_event_into(&mut buf) {
//                 Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gd" => {
//                     guides.push(ShapeGuide::load(e)?);
//                 }
//                 Ok(Event::End(ref e)) if e.local_name().as_ref() == b"avLst" => break,
//                 Ok(Event::Eof) => bail!("unexpected end of file."),
//                 Err(e) => bail!(e.to_string()),
//                 _ => (),
//             }
//         }
//         Ok(Self {
//             guide: Some(guides),
//         })
//     }
// }
