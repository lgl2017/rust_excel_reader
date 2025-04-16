use anyhow::bail;
use quick_xml::events::BytesStart;

use crate::{common_types::XSDDatetime, helper::string_to_datetime};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.dynamicfilter?view=openxml-3.0.1
///
/// This collection specifies dynamic filter criteria.
/// These criteria are considered dynamic because they can change, either with the data itself (e.g., "above average") or with the current system date (e.g., show values for "today").
/// For any cells whose values do not meet the specified criteria, the corresponding rows shall be hidden from view when the filter is applied.
///
/// Example
/// ```
/// <filterColumn colId="0">
///   <dynamicFilter type="today"/>
/// </filterColumn>
/// ```
///
/// dynamicFilter (Dynamic Filter)
#[derive(Debug, Clone, PartialEq)]
pub struct DynamicFilter {
    // Attributes
    /// maxValIso (Max ISO Value)
    ///
    /// A maximum value for dynamic filter.
    /// maxValIso shall be required for today, yesterday, tomorrow, nextWeek, thisWeek, lastWeek, nextMonth, thisMonth, lastMonth, nextQuarter, thisQuarter, lastQuarter, nextYear, thisYear, lastYear, and yearToDate.
    ///
    /// The above criteria are based on a value range;
    /// that is, if today's date is September 22nd, then the range for thisWeek is the values greater than or equal to September 17 and less than September 24.
    /// In the thisWeek range, the lower value is expressed valIso. The higher value is expressed using maxValIso.
    ///
    /// These dynamic filters shall not require valIso or maxValIso:
    /// Q1, Q2, Q3, Q4, M1, M2, M3, M4, M5, M6, M7, M8, M9, M10, M11 and M12.
    ///
    /// The above criteria shall not specify the range using valIso and maxValIso because Q1 always starts from M1 to M3, and M1 is always January.
    ///
    /// These types of dynamic filters shall use valIso and shall not use maxValIso:
    /// - aboveAverage and belowAverage
    pub max_val_iso: Option<XSDDatetime>,

    /// type (Dynamic filter type)
    ///
    /// Dynamic filter type, e.g., “today” or “nextWeek”.
    ///
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.dynamicfiltervalues?view=openxml-3.0.1
    pub filter_type: Option<String>,

    /// valIso (ISO Value)
    ///
    /// A minimum value for dynamic filter.
    pub min_val_iso: Option<XSDDatetime>,
}

impl DynamicFilter {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut filter = Self {
            max_val_iso: None,
            filter_type: None,
            min_val_iso: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"maxValIso" => filter.max_val_iso = string_to_datetime(&string_value),
                        b"type" => {
                            filter.filter_type = Some(string_value);
                        }
                        b"valIso" => {
                            filter.min_val_iso = string_to_datetime(&string_value);
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
