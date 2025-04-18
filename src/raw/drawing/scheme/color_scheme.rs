use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::excel::XmlReader;
use crate::raw::drawing::color::srgb_color::XlsxSrgbColor;
use crate::raw::drawing::color::system_color::XlsxSystemColor;
use crate::raw::drawing::color::XlsxColorEnum;

// use crate::raw::drawing::color::{
//     Accent1Color, Accent2Color, Accent3Color, Accent4Color, Accent5Color, Accent6Color, ColorEnum,
//     Dark1Color, Dark2Color, FollowedHyperlinkColor, HyperlinkColor, Light1Color, Light2Color,
//     SchemeColorEnum,
// };
use crate::common_types::HexColor;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorscheme?view=openxml-3.0.1
///
/// This element defines a set of colors which are referred to as a color scheme.
/// The color scheme is responsible for defining a list of twelve colors.
/// The twelve colors consist of six accent colors, two dark colors, two light colors and a color for each of a hyperlink and followed hyperlink.
///
/// Example:
/// ```
/// <clrScheme name="sample">
///     <dk1>
///         <sysClr val="windowText"/>
///     </dk1>
///     <lt1>
///         <sysClr val="window"/>
///     </lt1>
///     <dk2>
///         <srgbClr val="04617B"/>
///     </dk2>
///     <lt2>
///         <srgbClr val="DBF5F9"/>
///     </lt2>
///     <accent1>
///         <srgbClr val="0F6FC6"/>
///     </accent1>
///     <accent2>
///         <srgbClr val="009DD9"/>
///     </accent2>
///     <accent3>
///         <srgbClr val="0BD0D9"/>
///     </accent3>
///     <accent4>
///         <srgbClr val="10CF9B"/>
///     </accent4>
///     <accent5>
///         <srgbClr val="7CCA62"/>
///     </accent5>
///     <accent6>
///         <srgbClr val="A5C249"/>
///     </accent6>
///     <hlink>
///         <srgbClr val="FF9800"/>
///     </hlink>
///     <folHlink>
///         <srgbClr val="F45511"/>
///     </folHlink>
/// </clrScheme>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxColorScheme {
    // attribute
    pub name: Option<String>,

    // chilren
    /// Accent1Color: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.accent1color?view=openxml-3.0.1
    /// index: 4
    pub accent1: Option<HexColor>,

    /// Accent2Color: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.accent2color?view=openxml-3.0.1
    /// index: 5
    pub accent2: Option<HexColor>,

    /// Accent3Color: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.accent3color?view=openxml-3.0.1
    /// index: 6
    pub accent3: Option<HexColor>,

    /// Accent4Color: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.accent4color?view=openxml-3.0.1
    /// index: 7
    pub accent4: Option<HexColor>,

    /// Accent5Color: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.accent5color?view=openxml-3.0.1
    /// index: 8
    pub accent5: Option<HexColor>,

    /// Accent6Color: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.accent6color?view=openxml-3.0.1
    /// index: 9
    pub accent6: Option<HexColor>,

    /// Dark1Color: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.dark1color?view=openxml-3.0.1
    /// index: 0
    pub dk1: Option<HexColor>,

    /// Dark2Color: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.dark2color?view=openxml-3.0.1
    /// index: 2
    pub dk2: Option<HexColor>,

    /// FollowedHyperlinkColor: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.followedhyperlinkcolor?view=openxml-3.0.1
    /// index: 10
    // tag: folHlink
    pub fol_hlink: Option<HexColor>,

    /// HyperlinkColor: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.hyperlink?view=openxml-3.0.1
    /// index: 11
    pub hlink: Option<HexColor>,

    /// Light1Color: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.light1color?view=openxml-3.0.1
    /// index: 1
    pub lt1: Option<HexColor>,

    /// Light2Color: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.light2color?view=openxml-3.0.1
    /// index: 3
    pub lt2: Option<HexColor>,
}

/// Color used for defining clrScheme in themeElements.
///
/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorscheme?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
pub enum XlsxSchemeColorEnum {
    // srgbClr
    SrgbColor(XlsxSrgbColor),
    // sysClr
    SystemColor(XlsxSystemColor),
}

impl XlsxSchemeColorEnum {
    pub(crate) fn load(reader: &mut XmlReader, tag: &[u8]) -> anyhow::Result<Option<Self>> {
        let color_enum = XlsxColorEnum::load(reader, tag)?;
        if color_enum.is_none() {
            return Ok(None);
        }
        return match color_enum.unwrap() {
            XlsxColorEnum::SrgbColor(srgb_color) => return Ok(Some(Self::SrgbColor(srgb_color))),
            XlsxColorEnum::SystemColor(system_color) => {
                return Ok(Some(Self::SystemColor(system_color)))
            }
            _ => Ok(None),
        };
    }
}

impl XlsxSchemeColorEnum {
    pub(crate) fn to_hex(&self) -> Option<HexColor> {
        return match self {
            XlsxSchemeColorEnum::SrgbColor(srgb_color) => srgb_color.to_hex(),
            XlsxSchemeColorEnum::SystemColor(system_color) => system_color.to_hex(),
        };
    }
}

impl XlsxColorScheme {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();

        let mut color = Self {
            name: None,
            accent1: None,
            accent2: None,
            accent3: None,
            accent4: None,
            accent5: None,
            accent6: None,
            dk1: None,
            dk2: None,
            fol_hlink: None,
            hlink: None,
            lt1: None,
            lt2: None,
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

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"accent1" => {
                    if let Some(c) = XlsxSchemeColorEnum::load(reader, b"accent1")? {
                        color.accent1 = c.to_hex()
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"accent2" => {
                    if let Some(c) = XlsxSchemeColorEnum::load(reader, b"accent2")? {
                        color.accent2 = c.to_hex()
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"accent3" => {
                    if let Some(c) = XlsxSchemeColorEnum::load(reader, b"accent3")? {
                        color.accent3 = c.to_hex()
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"accent4" => {
                    if let Some(c) = XlsxSchemeColorEnum::load(reader, b"accent4")? {
                        color.accent4 = c.to_hex()
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"accent5" => {
                    if let Some(c) = XlsxSchemeColorEnum::load(reader, b"accent5")? {
                        color.accent5 = c.to_hex()
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"accent6" => {
                    if let Some(c) = XlsxSchemeColorEnum::load(reader, b"accent6")? {
                        color.accent6 = c.to_hex()
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dk1" => {
                    if let Some(c) = XlsxSchemeColorEnum::load(reader, b"dk1")? {
                        color.dk1 = c.to_hex()
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dk2" => {
                    if let Some(c) = XlsxSchemeColorEnum::load(reader, b"dk2")? {
                        color.dk2 = c.to_hex()
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"folHlink" => {
                    if let Some(c) = XlsxSchemeColorEnum::load(reader, b"folHlink")? {
                        color.fol_hlink = c.to_hex()
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"hlink" => {
                    if let Some(c) = XlsxSchemeColorEnum::load(reader, b"hlink")? {
                        color.hlink = c.to_hex()
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lt1" => {
                    if let Some(c) = XlsxSchemeColorEnum::load(reader, b"lt1")? {
                        color.lt1 = c.to_hex()
                    }
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"lt2" => {
                    if let Some(c) = XlsxSchemeColorEnum::load(reader, b"lt2")? {
                        color.lt2 = c.to_hex()
                    }
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"clrScheme" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(color)
    }
}

impl XlsxColorScheme {
    pub(crate) fn get_color(&self, index: u64) -> Option<HexColor> {
        return match index {
            0 => self.lt1.clone(),
            1 => self.dk1.clone(),
            2 => self.lt2.clone(),
            3 => self.dk2.clone(),
            4 => self.accent1.clone(),
            5 => self.accent2.clone(),
            6 => self.accent3.clone(),
            7 => self.accent4.clone(),
            8 => self.accent5.clone(),
            9 => self.accent6.clone(),
            10 => self.hlink.clone(),
            11 => self.fol_hlink.clone(),
            _ => None,
        };
    }
}
