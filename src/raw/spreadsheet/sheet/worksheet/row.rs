use std::io::Read;
use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_float, string_to_unsignedint},
};

use super::cell::XlsxCell;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.row?view=openxml-3.0.1
///
/// The element expresses information about an entire row of a worksheet,
/// and contains all cell definitions for a particular row in the worksheet.
///
/// Example:
/// ```
/// <row r="2" ht="20.25" customHeight="1">
///     <c r="A2" t="s" s="6">
///         <v>1</v>
///     </c>
///     <c r="B2" t="s" s="6">
///         <v>2</v>
///     </c>
///     <c r="C2" s="7" />
///     <c r="D2" s="8" />
///     <c r="E2">
///         <v>360</v>
///     </c>
/// </row>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxRow {
    /// extLst (Future Feature Data Storage Area) Not Supported

    /// Child Elements
    /// c (Cell)	§18.3.1.4
    pub cells: Option<Vec<XlsxCell>>,

    /// attributes

    /// collapsed (Collapsed)
    pub collapsed: Option<bool>,

    /// customFormat (Custom Format)
    pub custom_format: Option<bool>,

    /// customHeight (Custom Height)
    pub custom_height: Option<bool>,

    /// x14ac:dyDescent (DyDescent)
    ///
    /// https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/f11dfda4-46de-4035-8418-d76b0d3898f1
    ///
    /// specifies the vertical distance in pixels from the bottom of the cells in the current row to the typographical baseline of the cell content if, hypothetically, the zoom level for the sheet containing this row is 100 percent and the cell has bottom-alignment formatting.
    /// this property is only available in Office 2010 and later.
    ///
    /// Example:
    ///
    /// ```
    /// <row r="3" x14ac:dyDescent="0.25">
    /// ```
    pub dy_descent: Option<f64>,

    /// ht (Row Height)
    pub height: Option<f64>,

    /// hidden (Hidden)
    pub hidden: Option<bool>,

    /// outlineLevel (OutlineLevel)
    pub outline_level: Option<u64>,

    /// r (Row Index)
    pub row_index: Option<u64>,

    /// ph (Show Phonetic)
    pub show_phonetic: Option<bool>,

    /// spans (Spans)
    ///
    /// Example:
    /// ```
    /// // 1 row 11 cols
    /// <row r="3" spans="1:11" ht="55.05" customHeight="1" x14ac:dyDescent="0.25">
    /// ```
    pub spans: Option<(u64, u64)>,

    /// StyleIndex
    /// s (Style Index)
    pub style: Option<u64>,

    /// thickBot (Thick Bottom border)
    pub thick_bottom: Option<bool>,

    /// thickTop (Thick Top Border)
    pub thick_top: Option<bool>,
}

impl XlsxRow {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let mut row = Self {
            cells: None,
            collapsed: None,
            custom_format: None,
            custom_height: None,
            dy_descent: None,
            height: None,
            hidden: None,
            outline_level: None,
            row_index: None,
            show_phonetic: None,
            spans: None,
            style: None,
            thick_bottom: None,
            thick_top: None,
        };
        let mut cells: Vec<XlsxCell> = vec![];

        let attributes = e.attributes();
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"collapsed" => {
                            row.collapsed = string_to_bool(&string_value);
                        }
                        b"customFormat" => {
                            row.custom_format = string_to_bool(&string_value);
                        }
                        b"customHeight" => {
                            row.custom_height = string_to_bool(&string_value);
                        }
                        b"dyDescent" => {
                            row.dy_descent = string_to_float(&string_value);
                        }
                        b"ht" => {
                            row.height = string_to_float(&string_value);
                        }
                        b"hidden" => {
                            row.hidden = string_to_bool(&string_value);
                        }
                        b"outlineLevel" => {
                            row.outline_level = string_to_unsignedint(&string_value);
                        }
                        b"r" => {
                            row.row_index = string_to_unsignedint(&string_value);
                        }
                        b"ph" => {
                            row.show_phonetic = string_to_bool(&string_value);
                        }
                        b"spans" => {
                            let parts: Vec<&str> = string_value.split(|c| c == ':').collect();
                            if parts.len() == 2 {
                                let r = string_to_unsignedint(parts[0]);
                                let c = string_to_unsignedint(parts[1]);
                                if r.is_some() && c.is_some() {
                                    row.spans = Some((r.unwrap(), c.unwrap()))
                                }
                            }
                        }
                        b"s" => {
                            row.style = string_to_unsignedint(&string_value);
                        }
                        b"thickBot" => {
                            row.thick_bottom = string_to_bool(&string_value);
                        }
                        b"thickTop" => {
                            row.thick_top = string_to_bool(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        let mut buf: Vec<u8> = Vec::new();
        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"c" => {
                    cells.push(XlsxCell::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"row" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `row`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        if row.row_index.is_none() {
            bail!("row of unknwon index.")
        }

        row.cells = Some(cells);

        return Ok(row);
    }
}
