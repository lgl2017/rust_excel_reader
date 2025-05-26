use super::connection_site::XlsxConnectionSite;
use crate::excel::XmlReader;
use anyhow::bail;
use quick_xml::events::Event;
use std::io::Read;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.connectionsitelist?view=openxml-3.0.1
///
/// This element specifies all the connection sites that are used for this shape
///
/// Example
/// ```
/// <a:cxnLst>
///     <a:cxn ang="0">
///         <a:pos x="0" y="679622"/>
///     </a:cxn>
///     <a:cxn ang="0">
///         <a:pos x="1705233" y="679622"/>
///     </a:cxn>
/// </a:cxnLst>
/// ```
// tag: cxnLst
pub type XlsxConnectionSiteList = Vec<XlsxConnectionSite>;

pub(crate) fn load_connection_site_list(
    reader: &mut XmlReader<impl Read>,
) -> anyhow::Result<XlsxConnectionSiteList> {
    let mut buf = Vec::new();
    let mut sites: Vec<XlsxConnectionSite> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cxn" => {
                sites.push(XlsxConnectionSite::load(reader, e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cxnLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    Ok(sites)
}
