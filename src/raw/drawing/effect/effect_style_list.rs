use super::effect_style::XlsxEffectStyle;
use crate::excel::XmlReader;
use anyhow::bail;
use quick_xml::events::Event;

/// EffectStyleList: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.effectstylelist?view=openxml-3.0.1
///
/// Defines a set of three effect styles that create the effect style list for a theme.
/// The effect styles are arranged in order of subtle to moderate to intense.
///
/// Example:
/// ```
/// <effectStyleLst>
///   <effectStyle>
///     <effectLst>
///       <outerShdw blurRad="57150" dist="38100" dir="5400000" algn="ctr" rotWithShape="0">
/// …      </outerShdw>
///     </effectLst>
///   </effectStyle>
///   <effectStyle>
///     <effectLst>
///       <outerShdw blurRad="57150" dist="38100" dir="5400000" algn="ctr" rotWithShape="0">
/// …      </outerShdw>
///     </effectLst>
///   </effectStyle>
///   <effectStyle>
///     <effectLst>
///       <outerShdw blurRad="57150" dist="38100" dir="5400000" algn="ctr" rotWithShape="0">
/// …      </outerShdw>
///     </effectLst>
///     <scene3d>
/// …    </scene3d>
///     <sp3d prstMaterial="powder">
/// …    </sp3d>
///   </effectStyle>
/// </effectStyleLst>
/// ```
pub type XlsxEffectStyleList = Vec<XlsxEffectStyle>;

pub(crate) fn load_effect_style_list(
    reader: &mut XmlReader,
) -> anyhow::Result<XlsxEffectStyleList> {
    let mut styles: Vec<XlsxEffectStyle> = vec![];

    let mut buf = Vec::new();

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"effectStyle" => {
                styles.push(XlsxEffectStyle::load(reader)?);
            }

            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"effectStyleLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(styles)
}
