use std::io::Read;
use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;

use crate::raw::drawing::{
    effect::effect_style_list::{load_effect_style_list, XlsxEffectStyleList},
    fill::{
        load_bg_fill_style_lst, load_fill_style_lst, XlsxBackgroundFillStyleList, XlsxFillStyleList,
    },
    line::{load_line_style_list, XlsxLineStyleList},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.formatscheme?view=openxml-3.0.1
///
/// Example:
/// ```
/// <a:fmtScheme name="Office">
///     <a:fillStyleLst>
///         <a:solidFill>
///             <a:schemeClr val="phClr" />
///         </a:solidFill>
///         <a:gradFill rotWithShape="1">
///             <a:gsLst>
///                 <a:gs pos="0">
///                     <a:schemeClr val="phClr">
///                         <a:satMod val="103000" />
///                         <a:lumMod val="102000" />
///                         <a:tint val="94000" />
///                     </a:schemeClr>
///                 </a:gs>
///                 <a:gs pos="50000">
///                     <a:schemeClr val="phClr">
///                         <a:satMod val="110000" />
///                         <a:lumMod val="100000" />
///                         <a:shade val="100000" />
///                     </a:schemeClr>
///                 /a:gs>
///             </a:gsLst>
///             <a:lin ang="5400000" scaled="0" />
///         </a:gradFill>
///     </a:fillStyleLst>
///     <a:lnStyleLst>
///         <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
///             <a:solidFill>
///                 <a:schemeClr val="phClr" />
///             </a:solidFill>
///             <a:prstDash val="solid" />
///             <a:miter lim="800000" />
///         </a:ln>
///     </a:lnStyleLst>
///     <a:effectStyleLst>
///         <a:effectStyle>
///             <a:effectLst>
///                 <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
///                     <a:srgbClr val="000000">
///                         <a:alpha val="63000" />
///                     </a:srgbClr>
///                 </a:outerShdw>
///             </a:effectLst>
///         </a:effectStyle>
///     </a:effectStyleLst>
///     <a:bgFillStyleLst>
///         <a:solidFill>
///             <a:schemeClr val="phClr" />
///         </a:solidFill>
///     </a:bgFillStyleLst>
/// </a:fmtScheme>
/// ```
// tag: fmtScheme
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxFormatScheme {
    // attribute
    pub name: Option<String>,

    // children
    /// FillStyleList: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.fillstylelist?view=openxml-3.0.1
    /// This element defines a set of three fill styles that are used within a theme.
    /// The three fill styles are arranged in order from subtle to moderate to intense.
    // tag: fillStyleLst
    pub fill_style_lst: Option<XlsxFillStyleList>,

    /// BackgroundFillStyleList: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.backgroundfillstylelist?view=openxml-3.0.1
    /// This element defines a set of three background fills that are used within a theme.
    /// The three fill styles are arranged in order from subtle to moderate to intense.
    // tag: bgFillStyleLst
    pub bg_fill_style_lst: Option<XlsxBackgroundFillStyleList>,

    /// EffectStyleList: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectstylelist?view=openxml-3.0.1
    /// This element defines a set of three effect styles that create the effect style list for a theme.
    /// The effect styles are arranged in order of subtle to moderate to intense.
    // tag: effectStyleLst
    pub effect_style_lst: Option<XlsxEffectStyleList>,

    /// LineStyleList: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linestylelist?view=openxml-3.0.1
    /// defines a list of three line styles for use within a theme.
    /// The three line styles are arranged in order from subtle to moderate to intense versions of lines.
    // tag: lnStyleLst
    pub line_style_lst: Option<XlsxLineStyleList>,
}

impl XlsxFormatScheme {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut buf = Vec::new();
        let mut scheme = Self {
            name: None,
            fill_style_lst: None,
            bg_fill_style_lst: None,
            effect_style_lst: None,
            line_style_lst: None,
        };
        let attributes = e.attributes();
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"name" => {
                            scheme.name = Some(string_value);
                            break;
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fillStyleLst" => {
                    scheme.fill_style_lst = Some(load_fill_style_lst(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"bgFillStyleLst" => {
                    scheme.bg_fill_style_lst = Some(load_bg_fill_style_lst(reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectStyleLst" => {
                    let style_list = load_effect_style_list(reader)?;
                    scheme.effect_style_lst = Some(style_list);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lnStyleLst" => {
                    let style_list = load_line_style_list(reader)?;
                    scheme.line_style_lst = Some(style_list);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fmtScheme" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(scheme)
    }
}
