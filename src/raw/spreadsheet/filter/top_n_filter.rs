use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::helper::{string_to_bool, string_to_float};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.top10?view=openxml-3.0.1
///
/// This element specifies the top N (percent or number of items) to filter by.
///
/// Example:
/// ```
/// <filterColumn colId="0">
///   <top10 percent="1" val="5" filterVal="6"/>
/// </filterColumn
/// ```
///
/// top10 (Top 10)
#[derive(Debug, Clone, PartialEq)]
pub struct TopNFilter {
    // Attributes
    /// filterVal (Filter Value)
    ///
    /// The actual cell value in the range which is used to perform the comparison for this filter.
    pub filter_value: Option<f64>,

    /// percent (Filter by Percent)
    ///
    /// Flag indicating whether or not to filter by percent value of the column.
    /// A false value filters by number of items.
    pub filter_by_percent: Option<bool>,

    /// top (Top)
    ///
    /// Flag indicating whether or not to filter by top order.
    /// A false value filters by bottom order.
    pub filter_by_top: Option<bool>,

    /// val (Top or Bottom Value)
    ///
    /// Top or bottom value to use as the filter criteria.
    pub val: Option<f64>,
}
impl TopNFilter {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut filter = Self {
            filter_value: None,
            filter_by_percent: None,
            filter_by_top: None,
            val: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"filterVal" => filter.filter_value = string_to_float(&string_value),
                        b"percent" => filter.filter_by_percent = string_to_bool(&string_value),
                        b"top" => filter.filter_by_top = string_to_bool(&string_value),
                        b"val" => filter.val = string_to_float(&string_value),
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
