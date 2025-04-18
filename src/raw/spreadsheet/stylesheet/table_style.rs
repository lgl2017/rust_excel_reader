use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_int},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyles?view=openxml-3.0.1
///
/// This element represents a collection of Table style definitions for Table styles and PivotTable styles used in this workbook.
/// It consists of a sequence of tableStyle records, each defining a single Table style.
///
/// Example
/// ```
/// <tableStyles count="1" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">
///     <tableStyle name="Checklist" pivot="0" count="2" xr9:uid="{00000000-0011-0000-FFFF-FFFF00000000}">
///         <tableStyleElement type="wholeTable" dxfId="167" />
///         <tableStyleElement type="headerRow" dxfId="166" />
///     </tableStyle>
/// </tableStyles>
/// ```
// tag: tableStyles
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTableStyles {
    // children
    // tag: tableStyle
    pub table_style: Option<Vec<XlsxTableStyle>>,

    // attributes
    /// Name of the default table style to apply to new PivotTable
    // tag: defaultPivotStyle
    pub default_pivot_style: Option<String>,

    /// Name of default table style to apply to new Tables.
    // tag: defaultTableStyle
    pub default_table_style: Option<String>,
}

impl XlsxTableStyles {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut table_styles = Self {
            table_style: None,
            default_pivot_style: None,
            default_table_style: None,
        };

        let mut table_style: Vec<XlsxTableStyle> = vec![];

        let attributes = e.attributes();

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"defaultPivotStyle" => {
                            table_styles.default_pivot_style = Some(string_value);
                        }
                        b"defaultTableStyle" => {
                            table_styles.default_table_style = Some(string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tableStyle" => {
                    let style = XlsxTableStyle::load(reader, e)?;
                    table_style.push(style);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"tableStyles" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        table_styles.table_style = Some(table_style);

        Ok(table_styles)
    }
}

/// TableStyle: https:// - learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyle?view=openxml-3.0.1
// tag: tableStyle
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxTableStyle {
    // children
    /// [TableStyleElement]
    // tag: tableStyleElement
    pub table_style_element: Option<Vec<TableStyleElement>>,

    // attributes
    /// Name of this table style
    name: Option<String>,

    /// 'True' if this table style should be shown as an available pivot table style.
    pivot: Option<bool>,

    /// True if this table style should be shown as an available table style.
    table: Option<bool>,
}

impl XlsxTableStyle {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut style = Self {
            // children
            table_style_element: None,

            // attributes
            name: None,
            pivot: None,
            table: None,
        };

        let mut table_style_elements: Vec<TableStyleElement> = vec![];

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"name" => {
                            style.name = Some(string_value);
                        }
                        b"pivot" => {
                            style.pivot = string_to_bool(&string_value);
                        }
                        b"table" => {
                            style.table = string_to_bool(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tableStyleElement" => {
                    let style_element = TableStyleElement::load(e)?;
                    table_style_elements.push(style_element);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"tableStyle" => break,
                Ok(Event::Eof) => bail!("unexpected end of file."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        style.table_style_element = Some(table_style_elements);

        Ok(style)
    }
}

/// https:// - learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyleelement?view=openxml-3.0.1
/// Example:
/// ```
/// <tableStyleElement type="wholeTable" dxfId="167" />
/// <tableStyleElement type="headerRow" dxfId="166" />
/// ```
/// The order in which table style element formatting is applied is as follows:
/// - Table Style Element Order
/// - Whole Table
/// - First Column Stripe
/// - Second Column Stripe
/// - First Row Stripe
/// - Second Row Stripe
/// - Last Column
/// - First Column
/// - Header Row
/// - Total Row
/// - First Header Cell
/// - Last Header Cell
/// - First Total Cell
/// - Last Total Cell
/// For instance, row stripe formatting 'wins' over column stripe formatting, and both 'win' over whole table formatting.
/// - /
/// PivotTable Style Element Order
/// -  Whole Table
/// -  Page Field Labels
/// -  Page Field Values
/// -  First Column Stripe
/// -  Second Column Stripe
/// -  First Row Stripe
/// -  Second Row Stripe
/// -  First Column
/// -  Header Row
/// -  First Header Cell
/// -  Subtotal Column 1
/// -  Subtotal Column 2
/// -  Subtotal Column 3
/// -  Blank Row
/// -  Subtotal Row 1
/// -  Subtotal Row 2
/// -  Subtotal Row 3
/// -  Column Subheading 1
/// -  Column Subheading 2
/// -  Column Subheading 3
/// -  Row Subheading 1
/// -  Row Subheading 2
/// -  Row Subheading 3
/// -  Grand Total Column
/// -  Grand Total Row
// tag: tableStyleElement
#[derive(Debug, Clone, PartialEq)]
pub struct TableStyleElement {
    // attributes
    /// Zero-based index to a dxf record in the dxfs collection, specifying differential formatting to use with this Table or PivotTable style element.
    // tag: dxfId
    pub dxf_id: Option<i64>,

    /// Number of rows or columns in a single band of striping. Applies only when type is firstRowStripe, secondRowStripe, firstColumnStripe, or secondColumnStripe.
    pub size: Option<i64>,

    /// Identifies this table style element's type
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestylevalues?view=openxml-3.0.1
    pub r#type: Option<String>,
}

impl TableStyleElement {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut style = Self {
            dxf_id: None,
            size: None,
            r#type: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"dxfId" => style.dxf_id = string_to_int(&string_value),
                        b"size" => style.size = string_to_int(&string_value),
                        b"type" => style.r#type = Some(string_value),
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }

        Ok(style)
    }
}
