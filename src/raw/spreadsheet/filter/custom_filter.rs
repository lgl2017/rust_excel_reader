use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{excel::XmlReader, helper::string_to_bool};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.customfilters?view=openxml-3.0.1
///
/// When there is more than one custom filter criteria to apply (an 'and' or 'or' joining two criteria),
/// then this element groups the customFilter elements together.
///
/// There can be at most two customFilters specified.
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCustomFilters {
    // Child Elements	Subclause
    /// customFilter (Custom Filter Criteria)
    ///
    /// There can be at most two customFilters specified.
    pub custom_filter: Vec<XlsxCustomFilter>,

    // Attributes
    /// and (And)
    ///
    /// Flag indicating whether the two criterias have an "and" relationship.
    /// * '1'(true) indicates "and",
    /// * '0'(false) indicates "or".
    pub and: Option<bool>,
}

impl XlsxCustomFilters {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut and: Option<bool> = None;
        let mut filters: Vec<XlsxCustomFilter> = vec![];

        let attributes = e.attributes();
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"and" => {
                            and = string_to_bool(&string_value);
                            break;
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"customFilter" => {
                    filters.push(XlsxCustomFilter::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"customFilters" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `customFilters`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        if filters.len() > 2 {
            bail!("More than 2 custom filters found.")
        }

        return Ok(Self {
            custom_filter: filters,
            and,
        });
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.customfilter?view=openxml-3.0.1
///
/// A custom AutoFilter specifies an operator and a value.
///
/// There can be at most two customFilters specified, and in that case the parent element specifies whether the two conditions are joined by 'and' or 'or'.
/// For any cells whose values do not meet the specified criteria, the corresponding rows shall be hidden from view when the fitler is applied.
///
/// Example
/// ```
/// <customFilters and="1">
///   <customFilter operator="greaterThanOrEqual" val="0.2"/>
///   <customFilter operator="lessThanOrEqual" val="0.5"/>
/// </customFilters>
///
/// <customFilters>
///     <customFilter operator="greaterThan" val="0.5"/>
/// </customFilters>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxCustomFilter {
    // Attributes
    /// operator (Filter Comparison Operator)
    ///
    /// Operator used by the filter comparison.
    ///
    /// Possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.filteroperatorvalues?view=openxml-3.0.1
    pub operator: Option<String>,

    /// val (Top or Bottom Value)
    /// Top or bottom value used in the filter criteria.
    pub val: Option<String>,
}

impl XlsxCustomFilter {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut filter = Self {
            operator: None,
            val: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"operator" => filter.operator = Some(string_value),
                        b"val" => {
                            filter.val = Some(string_value);
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
