use super::XlsxColorEnum;
use crate::excel::XmlReader;

use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

///  CustomColorList: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.customcolorlist?view=openxml-3.0.1
///
///  Example:
///  ```
/// <a:custClrLst>
/// <a:custClr>
/// <a:sysClr val="windowText" lastClr="FFFFFF">
///      <a:alpha val="63000" />
///  </a:sysClr>
/// </a:custClr>
/// <a:custClr>
///  <a:srgbClr val="BCBCBC"/>
/// </a:custClr>
/// <a:custClr>
///  <a:scrgbClr r="50000" g="50000" b="50000"/>
/// </a:custClr>
/// <a:custClr>
///  <a:prstClr val="black" />
/// </a:custClr>
/// <a:custClr>
///  <a:schemeClr val="phClr">
///      <a:lumMod val="110000" />
///      <a:satMod val="105000" />
///      <a:tint val="67000" />
///  </a:schemeClr>
/// </a:custClr>
/// <a:custClr>
///  <a:srgbClr val="000000">
///      <a:alpha val="63000" />
///  </a:srgbClr>
/// </a:custClr>
/// <a:custClr>
///  <a:hslClr hue="14400000" sat="100.000%" lum="50.000%">
///      <a:alpha val="63000" />
///      <a:lumMod val="110000" />
///      <a:tint val="40000" />
///      <a:shade val="100000" />
///      <a:satMod val="350000" />
///      <a:comp/>
///      <a:inv/>
///  </a:hslClr>
/// </a:custClr>
/// </a:custClrLst>
///  ```
/// tag: custClrLst
pub type XlsxCustomColorList = Vec<XlsxCustomColor>;

pub(crate) fn load_custom_color_list(
    reader: &mut XmlReader,
) -> anyhow::Result<XlsxCustomColorList> {
    let mut buf = Vec::new();
    let mut colors: Vec<XlsxCustomColor> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"custClr" => {
                let color = XlsxCustomColor::load(reader, e)?;
                colors.push(color);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"custClrLst" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(colors)
}

///  CustomColor: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.customcolor?view=openxml-3.0.1
/// tag: custClr
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCustomColor {
    /// attributes
    pub name: Option<String>,

    /// children
    pub color: Option<XlsxColorEnum>,
}

impl XlsxCustomColor {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();

        let mut color = Self {
            name: None,
            color: XlsxColorEnum::load(reader, b"custClr")?,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"name" => {
                            color.name = Some(string_value);
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

        Ok(color)
    }
}
