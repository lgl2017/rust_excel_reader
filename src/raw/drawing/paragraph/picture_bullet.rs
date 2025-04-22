use crate::excel::XmlReader;
use crate::raw::drawing::image::blip::XlsxBlip;
use std::io::Read;
use anyhow::bail;
use quick_xml::events::Event;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.picturebullet?view=openxml-3.0.1
///
/// This element specifies that a picture be applied to a set of bullets.
///
/// Example
/// ```
/// <a:pPr …>
///     <a:buBlip>
///         <a:blip r:embed="rId2"/>
///     </a:buBlip>
/// </a:pPr>
/// ```
// tag: buBlip
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxPictureBullet {
    pub blip: Option<XlsxBlip>,
}

impl XlsxPictureBullet {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut bullet = Self { blip: None };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"blip" => {
                    bullet.blip = Some(XlsxBlip::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"blipFill" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(bullet)
    }
}
