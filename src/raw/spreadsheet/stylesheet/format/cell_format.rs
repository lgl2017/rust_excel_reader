use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_unsignedint},
};

use super::{alignment::XlsxAlignment, protection::XlsxCellProtection};

/// CellFormat: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformat?view=openxml-3.0.1
/// Example:
/// ```
/// // in cellStyleXfs
/// <cellStyleXfs count="4">
///     <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
///     <xf numFmtId="0" fontId="2" fillId="0" borderId="0" applyNumberFormat="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/>
///     <xf numFmtId="0" fontId="3" fillId="0" borderId="1" applyNumberFormat="0" applyFill="0" applyAlignment="0" applyProtection="0"/>
///     <xf numFmtId="0" fontId="4" fillId="2" borderId="2" applyNumberFormat="0" applyAlignment="0" applyProtection="0"/>
/// </cellStyleXfs>
/// // in cellXfs
/// <cellXfs count="1">
///     <xf numFmtId="0" fontId="0" applyNumberFormat="0" applyFont="1" applyFill="0" applyBorder="0" applyAlignment="1" applyProtection="0">
///         <alignment vertical="top" wrapText="1" />
///     </xf>
/// </cellXfs>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCellFormat {
    // children
    pub alignment: Option<XlsxAlignment>,
    pub protection: Option<XlsxCellProtection>,

    // attributes
    /// A boolean value indicating whether the alignment formatting specified for this xf should be applied.
    // tag: applyAlignment
    pub apply_alignment: Option<bool>,

    /// A boolean value indicating whether the border formatting specified for this xf should be applied.
    // tag: applyBorder
    pub apply_border: Option<bool>,

    /// A boolean value indicating whether the fill formatting specified for this xf should be applied.
    // tag: applyFill
    pub apply_fill: Option<bool>,

    /// A boolean value indicating whether the font formatting specified for this xf should be applied.
    // tag: applyFont
    pub apply_font: Option<bool>,

    /// A boolean value indicating whether the number formatting specified for this xf should be applied.
    // tag: applyNumberFormat
    pub apply_number_format: Option<bool>,

    /// A boolean value indicating whether the protection formatting specified for this xf should be applied.
    // tag: applyProtection
    pub apply_protection: Option<bool>,

    /// Zero-based index of the border record used by this cell format.
    // tag: borderId
    pub border_id: Option<u64>,

    /// Zero-based index of the fill record used by this cell format.
    // tag: fillId
    pub fill_id: Option<u64>,

    /// Zero-based index of the font record used by this cell format.
    // tag: fontId
    pub font_id: Option<u64>,

    /// Id of the number format (numFmt) record used by this cell format.
    // tag: numFmtId
    pub num_fmt_id: Option<u64>,

    /// A boolean value indicating whether the cell rendering includes a pivot table dropdown button.
    // tag: pivotButton
    pub pivot_button: Option<bool>,

    /// A boolean value indicating whether the text string in a cell should be prefixed by a single quote mark (e.g., 'text). In these cases, the quote is not stored in the Shared Strings Part.
    // tag: quotePrefix
    pub quote_prefix: Option<bool>,

    /// CellStyle.FormatId: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyle.formatid?view=openxml-3.0.1#documentformat-openxml-spreadsheet-cellstyle-formatid
    /// For xf records contained in cellXfs, this is the zero-based index of an xf record contained in cellStyleXfs corresponding to the cell style applied to the cell.
    /// Not present for xf records contained in cellStyleXfs.
    // tag: xfId
    pub xf_id: Option<u64>,
}

impl XlsxCellFormat {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut format = Self {
            // children
            alignment: None,
            protection: None,
            // attributes
            apply_alignment: None,
            apply_border: None,
            apply_fill: None,
            apply_font: None,
            apply_number_format: None,
            apply_protection: None,
            border_id: None,
            fill_id: None,
            font_id: None,
            num_fmt_id: None,
            pivot_button: None,
            quote_prefix: None,
            xf_id: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"applyAlignment" => {
                            format.apply_alignment = string_to_bool(&string_value);
                        }
                        b"applyBorder" => {
                            format.apply_border = string_to_bool(&string_value);
                        }
                        b"applyFill" => {
                            format.apply_fill = string_to_bool(&string_value);
                        }
                        b"applyFont" => {
                            format.apply_font = string_to_bool(&string_value);
                        }
                        b"applyNumberFormat" => {
                            format.apply_number_format = string_to_bool(&string_value);
                        }
                        b"applyProtection" => {
                            format.apply_protection = string_to_bool(&string_value);
                        }
                        b"borderId" => {
                            format.border_id = string_to_unsignedint(&string_value);
                        }
                        b"fillId" => {
                            format.fill_id = string_to_unsignedint(&string_value);
                        }
                        b"fontId" => {
                            format.font_id = string_to_unsignedint(&string_value);
                        }
                        b"numFmtId" => {
                            format.num_fmt_id = string_to_unsignedint(&string_value);
                        }
                        b"pivotButton" => {
                            format.pivot_button = string_to_bool(&string_value);
                        }
                        b"quotePrefix" => {
                            format.quote_prefix = string_to_bool(&string_value);
                        }
                        b"xfId" => {
                            format.xf_id = string_to_unsignedint(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"alignment" => {
                    let alignment = XlsxAlignment::load(e)?;
                    format.alignment = Some(alignment)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"protection" => {
                    let protection = XlsxCellProtection::load(e)?;
                    format.protection = Some(protection)
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"xf" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(format)
    }
}
