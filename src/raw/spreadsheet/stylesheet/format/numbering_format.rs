use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{excel::XmlReader, helper::string_to_unsignedint};

/// NumberingFormats: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformats?view=openxml-3.0.1
///
/// Example:
/// ```
/// <numFmts count="1">
///     <numFmt numFmtId="166" formatCode="General"/>
/// </numFmts>
/// ```
// tag: numFmts
pub type XlsxNumberingFormats = Vec<XlsxNumberingFormat>;

pub(crate) fn load_number_formats(reader: &mut XmlReader) -> anyhow::Result<XlsxNumberingFormats> {
    let mut buf: Vec<u8> = Vec::new();
    let mut formats: Vec<XlsxNumberingFormat> = vec![];

    loop {
        buf.clear();

        match reader.read_event_into(&mut buf) {
            Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"numFmt" => {
                let format = XlsxNumberingFormat::load(e)?;
                formats.push(format);
            }
            Ok(Event::End(ref e)) if e.local_name().as_ref() == b"numFmts" => break,
            Ok(Event::Eof) => bail!("unexpected end of file."),
            Err(e) => bail!(e.to_string()),
            _ => (),
        }
    }
    Ok(formats)
}

pub type FormatCode = String;

/// NumberingFormat: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-3.0.1
///
/// Example:
/// ```
/// <numFmt numFmtId="0" formatCode="General" />
/// ```
///
/// Note on "General" Format:
///
/// Alignment (Specified for Left-to-Right mode):
/// * Strings: left aligned
/// * Boolean/error values: centered
/// * Numbers: right aligned
/// * Dates: do not follow the "General" format, instead automatically convert to date formatting.
///
/// Numbers:
/// The application shall attempt to display the full number up to 11 digits (inc. decimal point).
/// If the number is too large, the application shall attempt to show exponential format.
/// If the number has too many significant digits, the display shall be truncated.
/// The optimal method of display is based on the available cell width.
/// If the number cannot be displayed using any of these formats in the available width, the application shall show "#" across the width of the cell.
///
/// Conditions for switching to exponential format:
/// The cell value shall have at least five digits for xE-xx.
/// If the exponent is bigger than the size allowed, a floating point number cannot fit, so try exponential notation.
/// Similarly, for negative exponents, check if there is space for even one (non-zero) digit in floating point format.
/// Finally, if there isn't room for all of the significant digits in floating point format (for a negative exponent), exponential format shall display more digits if the exponent is less than -3. (The 3 is because E-xx takes 4 characters, and the leading 0 in floating point takes only 1 character. Thus, for an exponent less than -3, there is more than 3 additional leading 0's, more than enough to compensate for the size of the E-xx.)
///
/// Floating point rule:
/// For general formatting in cells, max overall length for cell display is 11, not including negative sign, but includes leading zeros and decimal separator.
///
/// tag: numFmt
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxNumberingFormat {
    // attributes
    /// The number format code for this number format.
    // tag: formatCode
    pub format_code: Option<FormatCode>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat.numberformatid?view=openxml-3.0.1#documentformat-openxml-spreadsheet-numberingformat-numberformatid
    /// Id used by the master style records (xf's) to reference this number format.
    // tag: numFmtId
    pub num_fmt_id: Option<u64>,
}

impl XlsxNumberingFormat {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut format = Self {
            format_code: None,
            num_fmt_id: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"formatCode" => format.format_code = Some(string_value),
                        b"numFmtId" => format.num_fmt_id = string_to_unsignedint(&string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(format)
    }
}

pub(crate) fn get_builtin_format_code(number_format_id: u64) -> Option<FormatCode> {
    let str = match number_format_id {
        0 => "general",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        11 => "0.00E+00",
        12 => "# ?/?",
        13 => "# ??/??",
        14 => "mm-dd-yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        18 => "h =>mm AM/PM",
        19 => "h =>mm =>ss AM/PM",
        20 => "hh =>mm",
        21 => "hh =>mm =>ss",
        22 => "m/d/yy hh =>mm",
        37 => "#,##0 ;(#,##0)",
        38 => "#,##0 ;[red](#,##0)",
        39 => "#,##0.00 ;(#,##0.00)",
        40 => "#,##0.00 ;[red](#,##0.00)",
        41 => "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)",
        42 => "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)",
        43 => "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)",
        44 => "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)",
        45 => "mm =>ss",
        46 => "[h] =>mm =>ss",
        47 => "mm =>ss.0",
        48 => "##0.0E+0",
        49 => "@",
        _ => "",
    };
    return if str.is_empty() {
        None
    } else {
        Some(str.to_string())
    };
}
