use super::tab_stop::XlsxTabStop;
use crate::excel::XmlReader;
use std::io::Read;
use anyhow::bail;
use quick_xml::events::Event;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tabstoplist?view=openxml-3.0.1
///
/// This element specifies the list of all tab stops that are to be used within a paragraph.
/// These tabs should be used when describing any custom tab stops within the document.
/// If these are not specified then the default tab stops of the generating application should be used.
///
/// Example:
/// ```
/// <a:tabLst>
///     <a:tab pos="2292350" algn="l"/>
///     <a:tab pos="2627313" algn="l"/>
///     <a:tab pos="2743200" algn="l"/>
///     <a:tab pos="2974975" algn="l"/>
/// </a:tabLst>
/// ```
// tag: tabLst
pub type XlsxTabStopList = Vec<XlsxTabStop>;

pub(crate) fn load_tab_stop_list(reader: &mut XmlReader<impl Read>) -> anyhow::Result<XlsxTabStopList> {
    let mut stops: Vec<XlsxTabStop> = vec![];
    let mut buf = Vec::new();

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) => {
                stops.push(XlsxTabStop::load(e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"tabLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(stops)
}
