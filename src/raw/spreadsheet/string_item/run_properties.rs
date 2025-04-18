use anyhow::bail;
use quick_xml::events::Event;

use crate::{
    excel::XmlReader,
    helper::{extract_val_attribute, string_to_bool, string_to_float, string_to_unsignedint},
    raw::spreadsheet::stylesheet::color::XlsxColor,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.runproperties?view=openxml-3.0.1
///
/// This element represents a set of properties to apply to the contents of this rich text run.
///
/// Example:
/// ```
/// <r>
///     <rPr>
///         <i val="1" />
///         <sz val="10" />
///         <color indexed="8" />
///         <rFont val="Helvetica Neue" />
///     </rPr>
///     <t>italic</t>
/// </r>
/// ```
/// rPr (Rich Text Run)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxRunProperties {
    // children
    /// Bold: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.bold?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <b val="0"/>
    /// <b /> // If tag presented and val omitted, the default value is true.
    /// ```
    pub bold: Option<bool>,

    /// FontCharset: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontcharset?view=openxml-3.0.1
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

    /// Outline: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.outline?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <outline val="0"/>
    /// <outline /> // If tag presented and val omitted, the default value is true.
    /// ```
    pub outline: Option<bool>,

    /// RunFont: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.runfont?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <rFont val="Helvetica Neue"/>
    /// ```
    // rFont (Font)
    pub run_font: Option<String>,

    /// FontScheme: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontscheme?view=openxml-3.0.1
    ///
    /// Example:
    /// ```
    /// <scheme val="minor"/>
    /// ```
    /// Allowed values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fontschemevalues?view=openxml-3.0.1
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

impl XlsxRunProperties {
    pub(crate) fn load(reader: &mut XmlReader) -> anyhow::Result<Self> {
        let mut buf = Vec::new();

        let mut properties = Self {
            bold: None,
            charset: None,
            color: None,
            condense: None,
            extend: None,
            family: None,
            italic: None,
            outline: None,
            run_font: None,
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
                    properties.bold = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"charset" => {
                    properties.charset = extract_val_attribute(e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"color" => {
                    properties.color = Some(XlsxColor::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"condense" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    properties.condense = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extend" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    properties.extend = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"family" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    properties.family = string_to_unsignedint(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"i" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    properties.italic = string_to_bool(&val_string);
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"outline" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    properties.outline = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"rFont" => {
                    properties.run_font = extract_val_attribute(e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"scheme" => {
                    properties.scheme = extract_val_attribute(e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"shadow" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    properties.shadow = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"strike" => {
                    let val_string = extract_val_attribute(e)?.unwrap_or("1".to_owned());
                    properties.strike = string_to_bool(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sz" => {
                    let Some(val_string) = extract_val_attribute(e)? else {
                        continue;
                    };
                    properties.size = string_to_float(&val_string);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"u" => {
                    properties.underline = extract_val_attribute(e)?;
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"vertAlign" => {
                    properties.vert_align = extract_val_attribute(e)?;
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"rPr" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(properties)
    }
}
