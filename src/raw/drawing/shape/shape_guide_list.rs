use anyhow::bail;
use quick_xml::events::Event;
use crate::excel::XmlReader;
use super::shape_guide::ShapeGuide;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapeguidelist?view=openxml-3.0.1
///
/// This element specifies all the guides that are used for this shape
///
/// Example:
/// ```
/// <a:gdLst>
///     <a:gd name="myGuide" fmla="*/ h 2 3"/>
/// </a:gdLst>
/// ```

pub type ShapeGuideList = Vec<ShapeGuide>;

pub(crate) fn load_shape_guide_list(reader: &mut XmlReader) -> anyhow::Result<ShapeGuideList> {
    let mut buf = Vec::new();
    let mut guides: Vec<ShapeGuide> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gd" => {
                guides.push(ShapeGuide::load(e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"gdLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    Ok(guides)
}

// #[derive(Debug, Clone, PartialEq)]
// pub struct ShapeGuideList {
//     // children
//     // gd (Shape Guide)
//     pub guide: Option<Vec<ShapeGuide>>,
// }

// impl ShapeGuideList {
//     pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
//         let mut buf = Vec::new();
//         let mut guides: Vec<ShapeGuide> = vec![];

//         loop {
//             buf.clear();

//             match reader.read_event_into(&mut buf) {
//                 Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gd" => {
//                     guides.push(ShapeGuide::load(e)?);
//                 }
//                 Ok(Event::End(ref e)) if e.local_name().as_ref() == b"gdLst" => break,
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
