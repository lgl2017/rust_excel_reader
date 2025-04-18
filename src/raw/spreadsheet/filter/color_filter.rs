use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::{string_to_bool, string_to_unsignedint};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.colorfilter?view=openxml-3.0.1
///
/// This element specifies the color to filter by and whether to use the cell's fill or font color in the filter criteria.
/// If the cell's font or fill color does not match the color specified in the criteria, the rows corresponding to those cells are hidden from view.
///
/// Example
/// ```
/// <filterColumn colId="1">
///   <colorFilter dxfId="0" cellColor="0"/>
/// </filterColumn>
/// ```
///
/// colorFilter (Color Filter Criteria)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxColorFilter {
    // Attributes
    /// cellColor (Filter By Cell Color)
    ///
    /// Flag indicating whether or not to filter by the cell's fill color.
    ///
    /// * '1'(true): indicates to filter by cell fill.
    /// * '0'(false): indicates to filter by the cell's font color.
    ///
    /// For rich text in cells, if the color specified appears in the cell at all, it shall be included in the filter.
    pub cell_color: Option<bool>,

    /// dxfId (Differential Format Record Id)
    ///
    /// Id of differential format record (dxf) in the Styles Part which expresses the color value to filter by.
    pub dxf_id: Option<u64>,
}

impl XlsxColorFilter {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut filter = Self {
            cell_color: None,
            dxf_id: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"cellColor" => filter.cell_color = string_to_bool(&string_value),
                        b"dxfId" => {
                            filter.dxf_id = string_to_unsignedint(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        Ok(filter)
    }
}
