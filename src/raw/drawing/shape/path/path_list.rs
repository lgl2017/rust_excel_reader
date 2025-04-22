use super::XlsxPath;
use crate::excel::XmlReader;
use std::io::Read;
use anyhow::bail;
use quick_xml::events::Event;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.pathlist?view=openxml-3.0.1
///
/// This element specifies the entire path that is to make up a single geometric shape.
///
/// Example
/// ```
///   <a:pathLst>
///     <a:path w="2650602" h="1261641">
///       <a:moveTo>
///         <a:pt x="0" y="1261641"/>
///       </a:moveTo>
///       <a:lnTo>
///         <a:pt x="2650602" y="1261641"/>
///       </a:lnTo>
///       <a:lnTo>
///         <a:pt x="1226916" y="0"/>
///       </a:lnTo>
///       <a:close/>
///     </a:path>
///   </a:pathLst>
/// ```
pub type XlsxPathList = Vec<XlsxPath>;

pub(crate) fn load_path_list(reader: &mut XmlReader<impl Read>) -> anyhow::Result<XlsxPathList> {
    let mut buf = Vec::new();
    let mut paths: Vec<XlsxPath> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"path" => {
                paths.push(XlsxPath::load(reader, e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"pathLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    Ok(paths)
}
