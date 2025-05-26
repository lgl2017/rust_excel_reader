use super::shape_guide::XlsxShapeGuide;
use crate::excel::XmlReader;
use anyhow::bail;
use quick_xml::events::Event;
use std::io::Read;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapeguidelist?view=openxml-3.0.1
///
/// This element specifies all the guides that are used for this shape
///
/// A guide is specified by the gd element and defines a calculated value that can be used for the construction of the corresponding shape.
/// Guides that have a literal value formula specified via fmla="val x" should only be used within the avLst as an adjust value for the shape. This however is not strictly enforced.
///
///
/// Example:
/// ```
/// <a:custGeom>
///   <a:avLst/>
///   <a:gdLst>
///     <a:gd name="myGuide" fmla="*/ h 2 3"/>
///   </a:gdLst>
///   <a:ahLst/>
///   <a:cxnLst/>
///   <a:rect l="0" t="0" r="0" b="0"/>
///   <a:pathLst>
///     <a:path w="1705233" h="679622">
///       <a:moveTo>
///         <a:pt x="0" y="myGuide"/>
///       </a:moveTo>
///       <a:lnTo>
///         <a:pt x="1705233" y="myGuide"/>
///       </a:lnTo>
///       <a:lnTo>
///         <a:pt x="852616" y="0"/>
///       </a:lnTo>
///       <a:close/>
///     </a:path>
///   </a:pathLst>
/// </a:custGeom>
/// ```
/// gdLst (Shape guide list)
pub type XlsxShapeGuideList = Vec<XlsxShapeGuide>;

pub(crate) fn load_shape_guide_list(
    reader: &mut XmlReader<impl Read>,
) -> anyhow::Result<XlsxShapeGuideList> {
    let mut buf = Vec::new();
    let mut guides: Vec<XlsxShapeGuide> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"gd" => {
                guides.push(XlsxShapeGuide::load(e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"gdLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    Ok(guides)
}
