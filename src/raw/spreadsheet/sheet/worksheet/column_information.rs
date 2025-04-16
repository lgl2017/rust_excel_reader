use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_float, string_to_unsignedint},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.columns?view=openxml-3.0.1
///
/// Information about whole columns of the worksheet.
///
/// Eample:
/// ```
/// <cols>
///   <col min="4" max="4" width="12" bestFit="1" customWidth="1"/>
///   <col min="5" max="5" width="9.140625" style="3"/>
/// </cols>
/// ```
pub type ColumnInformations = Vec<ColumnInformation>;

pub(crate) fn load_column_infos(reader: &mut XmlReader) -> anyhow::Result<ColumnInformations> {
    let mut cols: ColumnInformations = vec![];

    let mut buf = Vec::new();
    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"col" => {
                cols.push(ColumnInformation::load(e)?);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"cols" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }

    Ok(cols)
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column?view=openxml-3.0.1
///
/// Defines column width and column formatting for one or more columns of the worksheet.
///
/// Example
/// ```
/// <col min="5" max="5" width="9.140625" style="3"/>
/// <col min="1" max="5" width="16.3516" style="1" customWidth="1" />
/// ```
/// col (Column Width & Formatting)
#[derive(Debug, Clone, PartialEq)]
pub struct ColumnInformation {
    /// Attributes
    /// bestFit (Best Fit Column Width)
    ///
    /// Flag indicating if the specified column(s) is set to 'best fit'.
    /// 'Best fit' is set to true under these conditions:
    /// - The column width has never been manually set by the user,
    /// - The column width is not the default width
    /// - 'Best fit' means that when numbers are typed into a cell contained in a 'best fit' column, the column width should automatically resize to display the number.
    pub best_fit: Option<bool>,

    /// collapsed (Collapsed)
    ///
    /// Flag indicating if the outlining of the affected column(s) is in the collapsed state.
    pub collapsed: Option<bool>,

    /// customWidth (Custom Width)
    ///
    /// Flag indicating that the column width for the affected column(s) is different from the default or has been manually set.
    pub custom_width: Option<bool>,

    /// hidden (Hidden Columns)
    ///
    /// Flag indicating if the affected column(s) are hidden on this worksheet.
    pub hidden: Option<bool>,

    /// max (Maximum Column)
    ///
    /// Last column affected by this 'column info' record.
    pub max_column: Option<u64>,

    /// min (Minimum Column)
    ///
    /// First column affected by this 'column info' record.
    pub min_column: Option<u64>,

    /// outlineLevel (Outline Level)
    ///
    /// Outline level of affected column(s).
    /// Range is 0 to 7.
    ///
    /// unsignedByte
    pub outline_level: Option<u64>,

    /// phonetic (Show Phonetic Information)
    ///
    /// Flag indicating if the phonetic information should be displayed by default for the affected column(s) of the worksheet.
    pub show_phonetic: Option<bool>,

    /// style (Style)
    ///
    /// Default style for the affected column(s).
    /// Affects cells not yet allocated in the column(s).
    /// In other words, this style applies to new columns.
    ///
    /// 0 based index reference to `cellXfs` in stylesheet.
    pub style: Option<u64>,

    /// width (Column Width)
    ///
    /// Column width measured as the number of characters of the maximum digit width of the numbers 0, 1, 2, …, 9 as rendered in the normal style's font.
    /// There are 4 pixels of margin padding (two on each side), plus 1 pixel padding for the gridlines.
    ///
    /// value of this attirbute from Number of Characters:
    /// ```
    /// width = Truncate([{Number of Characters} * {Maximum Digit Width} + {5 pixel padding}]/{Maximum Digit Width}*256)/256
    /// ```
    ///
    /// To translate the value of this attribute into the column width value at runtime (expressed in terms of pixels):
    /// ```
    /// =Truncate(((256 * {width} + Truncate(128/{Maximum Digit Width}))/256)*{Maximum Digit Width})
    /// ```
    ///
    /// To translate from pixels to character width"
    /// ```
    /// =Truncate(({pixels}-5)/{Maximum Digit Width} * 100+0.5)/100
    /// ```
    pub width: Option<f64>,
}

impl ColumnInformation {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut info = Self {
            best_fit: None,
            collapsed: None,
            custom_width: None,
            hidden: None,
            max_column: None,
            min_column: None,
            outline_level: None,
            show_phonetic: None,
            style: None,
            width: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"bestFit" => {
                            info.best_fit = string_to_bool(&string_value);
                        }
                        b"collapsed" => {
                            info.collapsed = string_to_bool(&string_value);
                        }
                        b"customWidth" => {
                            info.custom_width = string_to_bool(&string_value);
                        }
                        b"hidden" => {
                            info.hidden = string_to_bool(&string_value);
                        }
                        b"max" => {
                            info.max_column = string_to_unsignedint(&string_value);
                        }
                        b"min" => {
                            info.min_column = string_to_unsignedint(&string_value);
                        }
                        b"outlineLevel" => {
                            info.outline_level = string_to_unsignedint(&string_value);
                        }
                        b"phonetic" => {
                            info.show_phonetic = string_to_bool(&string_value);
                        }
                        b"style" => {
                            info.style = string_to_unsignedint(&string_value);
                        }
                        b"width" => {
                            info.width = string_to_float(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        Ok(info)
    }
}
