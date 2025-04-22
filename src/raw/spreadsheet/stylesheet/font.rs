use anyhow::bail;
use quick_xml::events::Event;
use std::io::Read;

use crate::{
    excel::XmlReader,
    helper::{extract_val_attribute, string_to_bool, string_to_float, string_to_unsignedint},
};

use super::color::XlsxColor;

/// Fonts: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fonts?view=openxml-3.0.1
pub type XlsxFonts = Vec<XlsxFont>;

pub(crate) fn load_fonts(reader: &mut XmlReader<impl Read>) -> anyhow::Result<XlsxFonts> {
    let mut buf = Vec::new();
    let mut fonts: Vec<XlsxFont> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"font" => {
                let font = XlsxFont::load(reader)?;
                fonts.push(font);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"fonts" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(fonts)
}

/// xml tag: font
/// Font: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.font?view=openxml-3.0.1
///
/// Example:
/// ```
/// <fonts count="2">
///     <font>
///         <sz val="11"/>
///         <color theme="1"/>
///         <name val="Calibri"/>
///         <family val="2"/>
///         <scheme val="minor"/>
///     </font>
///     <font>
///         <strike/>
///         <sz val="12"/>
///         <color theme="1"/>
///         <name val="Arial"/>
///         <family val="2"/>
///     </font>
/// </fonts>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxFont {
    // children
    /// Bold: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.bold?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <b val="0"/>
    /// <b /> // If tag presented and val omitted, the default value is true.
    /// ```
    pub bold: Option<bool>,

    /// FontCharset: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontcharset?view=openxml-3.0.1.
    ///
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.linq.x.charset?view=openxml-3.0.1#documentformat-openxml-linq-x-charset
    ///
    /// Example:
    /// ```
    /// <charset val="..."/>
    /// ```
    pub charset: Option<String>,

    ///  color
    pub color: Option<XlsxColor>,

    /// Condense: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.condense?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <condense val="0"/>
    /// <condense /> // If tag presented and val omitted, the default value is true.
    /// ```
    pub condense: Option<bool>,

    /// Extend: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.extend?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <extend val="0"/>
    /// <extend /> // If tag presented and val omitted, the default value is true.
    /// ```
    pub extend: Option<bool>,

    /// FontFamily: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontfamily?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <family val="2"/>
    /// ```
    ///
    /// 0: Not applicable.
    /// 1: Roman
    /// 2: Swiss
    /// 3: Modern
    /// 4: Script
    /// 5: Decorative
    ///
    /// can be converted to string using the provided [get_family_name] method.
    pub family: Option<u64>,

    /// Italic: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.italic?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <i val="0"/>
    /// <i /> // If tag presented and val omitted, the default value is true.
    /// ```
    pub italic: Option<bool>,

    /// FontName: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontname?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <name val="Arial"/>
    /// ```
    pub name: Option<String>,

    /// Outline: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.outline?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <outline val="0"/>
    /// <outline /> // If tag presented and val omitted, the default value is true.
    /// ```
    pub outline: Option<bool>,

    /// FontScheme: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontscheme?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <scheme val="minor"/>
    /// ```
    /// Allowed values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontschemevalues?view=openxml-3.0.1
    ///
    /// major: Major font
    /// minor: Minor font
    /// none: None
    pub scheme: Option<String>,

    /// Shadow: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.shadow?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <shadow val="0"/>
    /// <shadow /> // If tag presented and val omitted, the default value is true.
    /// ```
    pub shadow: Option<bool>,

    /// Strike: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.strike?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <strike val="0"/>
    /// <strike /> // If tag presented and val omitted, the default value is true.
    /// ```
    pub strike: Option<bool>,

    /// FontSize: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontsize?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <sz val="11"/>
    /// ```
    pub size: Option<f64>,

    /// Underline: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.underline?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <u val="double"/>
    /// ```
    /// Allowed Values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.underlinevalues?view=openxml-3.0.1
    pub underline: Option<String>,

    /// VerticalTextAlignment: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.verticaltextalignment?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <vertAlign val="subscript"/>
    /// ```
    /// Allowed Values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.verticalalignmentrunvalues?view=openxml-3.0.1
    // xml tag: vertAlign
    pub vert_align: Option<String>,
}

impl XlsxFont {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>) -> anyhow::Result<Self> {
        let mut buf = Vec::new();

        let mut font = Self {
            bold: None,
            charset: None,
            color: None,
            condense: None,
            extend: None,
            family: None,
            italic: None,
            name: None,
            outline: None,
            scheme: None,
            shadow: None,
            strike: None,
            size: None,
            underline: None,
            vert_align: None,
        };

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"b" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    font.bold = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"charset" => {
                    font.charset = extract_val_attribute(e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"color" => {
                    font.color = Some(XlsxColor::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"condense" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    font.condense = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extend" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    font.extend = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"family" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    font.family = string_to_unsignedint(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"i" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    font.italic = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"name" => {
                    font.name = extract_val_attribute(e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"outline" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    font.outline = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"scheme" => {
                    font.scheme = extract_val_attribute(e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"shadow" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    font.shadow = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"strike" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    font.strike = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sz" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    font.size = string_to_float(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"u" => {
                    font.underline = extract_val_attribute(e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"vertAlign" => {
                    font.vert_align = extract_val_attribute(e)?;
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"font" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(font)
    }
}
