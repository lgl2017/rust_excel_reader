use crate::excel::XmlReader;
use crate::raw::drawing::color::color_map::XlsxColorMap;
use anyhow::bail;
use quick_xml::events::Event;

use super::color_scheme::XlsxColorScheme;

/// ExtraColorSchemeList: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.extracolorschemelist?view=openxml-3.0.1
// tag: extraClrSchemeLst
pub type XlsxExtraColorSchemeList = Vec<XlsxExtraColorScheme>;

pub(crate) fn load_extra_color_scheme_list(
    reader: &mut XmlReader,
) -> anyhow::Result<Vec<XlsxExtraColorScheme>> {
    let mut buf = Vec::new();
    let mut schemes: Vec<XlsxExtraColorScheme> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extraClrScheme" => {
                let scheme = XlsxExtraColorScheme::load(reader)?;
                schemes.push(scheme);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"extraClrSchemeLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(schemes)
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.extracolorscheme?view=openxml-3.0.1
///
/// Example:
/// ```
/// <extraClrScheme>
///     <clrScheme name="sample">
///         <dk1>
///             <sysClr val="windowText"/>
///         </dk1>
///         <lt1>
///             <sysClr val="window"/>
///         </lt1>
///         <dk2>
///             <srgbClr val="04617B"/>
///         </dk2>
///         <lt2>
///             <srgbClr val="DBF5F9"/>
///         </lt2>
///         <accent1>
///             <srgbClr val="0F6FC6"/>
///         </accent1>
///         <accent2>
///             <srgbClr val="009DD9"/>
///         </accent2>
///         <accent3>
///             <srgbClr val="0BD0D9"/>
///         </accent3>
///         <accent4>
///             <srgbClr val="10CF9B"/>
///         </accent4>
///         <accent5>
///             <srgbClr val="7CCA62"/>
///         </accent5>
///         <accent6>
///             <srgbClr val="A5C249"/>
///         </accent6>
///         <hlink>
///             <srgbClr val="FF9800"/>
///         </hlink>
///         <folHlink>
///             <srgbClr val="F45511"/>
///         </folHlink>
///     </clrScheme>
///     <clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1"   accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5"  accent6="accent6" hlink="hlink" folHlink="folHlink"/>
/// </extraClrScheme>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxExtraColorScheme {
    // children
    // tag: clrScheme
    pub color_scheme: Option<XlsxColorScheme>,

    // tag: clrMap
    pub color_map: Option<XlsxColorMap>,
}

impl XlsxExtraColorScheme {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut scheme = Self {
            color_scheme: None,
            color_map: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrScheme" => {
                    let color_scheme = XlsxColorScheme::load(reader, e)?;
                    scheme.color_scheme = Some(color_scheme);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"clrMap" => {
                    let map = XlsxColorMap::load(e)?;
                    scheme.color_map = Some(map);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"extraClrScheme" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(scheme)
    }
}
