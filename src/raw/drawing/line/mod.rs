use crate::excel::XmlReader;

use anyhow::bail;
use std::io::Read;

use outline::XlsxOutline;
use quick_xml::events::Event;

pub mod custom_dash;
pub mod head_end;
pub mod line_join_bevel;
pub mod line_reference;
pub mod miter;
pub mod outline;
pub mod round_line_join;
pub mod tail_end;

/// XlsxLineStyleList: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linestylelist?view=openxml-3.0.1
/// defines a list of three line styles for use within a theme.
/// The three line styles are arranged in order from subtle to moderate to intense versions of lines.
///
/// Example:
/// ```
/// <lnStyleLst>
///   <ln w="9525" cap="flat" cmpd="sng" algn="ctr">
///     <solidFill>
///       <schemeClr val="phClr">
///         <shade val="50000"/>
///         <satMod val="103000"/>
///       </schemeClr>
///     </solidFill>
///     <prstDash val="solid"/>
///   </ln>
///   <ln w="25400" cap="flat" cmpd="sng" algn="ctr">
///     <solidFill>
///       <schemeClr val="phClr"/>
///     </solidFill>
///     <prstDash val="solid"/>
///   </ln>
///   <ln w="38100" cap="flat" cmpd="sng" algn="ctr">
///     <solidFill>
///       <schemeClr val="phClr"/>
///     </solidFill>
///     <prstDash val="solid"/>
///   </ln>
/// </lnStyleLst>
/// ```
pub type XlsxLineStyleList = Vec<XlsxOutline>;

pub(crate) fn load_line_style_list(
    reader: &mut XmlReader<impl Read>,
) -> anyhow::Result<XlsxLineStyleList> {
    let mut outlines: Vec<XlsxOutline> = vec![];

    let mut buf = Vec::new();

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"ln" => {
                outlines.push(XlsxOutline::load(reader, e)?);
            }

            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"lnStyleLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(outlines)
}
