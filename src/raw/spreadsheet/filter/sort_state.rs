use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    common_types::Dimension,
    excel::XmlReader,
    helper::{string_to_bool, string_to_unsignedint},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sortstate?view=openxml-3.0.1
///
/// This collection preserves the AutoFilter sort state.
///
/// Example:
/// This example shows a sort which is case-sensitive, descending sort.
/// While the range of data to sort is B4:E8, the range to sort by is B4:B8.
/// ```
/// <sortState caseSensitive="1" ref="B4:E8">
///     <sortCondition descending="1" ref="B4:B8"/>
/// </sortState>
/// ```
///
/// sortState (Sort State)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxSortState {
    // extLst (Future Feature Data Storage Area) Not supported

    // Child Elements
    /// sortCondition (Sort Condition)
    ///
    /// When more than one sortCondition is specified, the first condition is applied first, then the second condition is applied, and so on.
    pub sort_condition: Option<Vec<XlsxSortCondition>>,

    // Attributes
    /// caseSensitive (Case Sensitive)
    ///
    /// Flag indicating whether or not the sort is case-sensitive.
    pub case_sensitive: Option<bool>,

    /// columnSort (Sort by Columns)
    ///
    /// Flag indicating whether or not to sort by columns.
    /// Only applies to ranges that don’t have AutoFilter applied.
    pub column_sort: Option<bool>,

    /// ref (Sort Range)
    ///
    /// The whole range of data to sort (not just the sort-by column).
    pub r#ref: Option<Dimension>,

    /// sortMethod (Sort Method)
    ///
    /// Strokes or PinYin sort method.
    /// Applies only to these application UI languages:
    /// - Chinese Simplified
    /// - Chinese Traditional
    /// - Japanese
    ///  For these languages, alternate sort methods can be selected, affecting how the data is sorted.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sortmethodvalues?view=openxml-3.0.1
    pub sort_method: Option<String>,
}

impl XlsxSortState {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut sort_state = Self {
            sort_condition: None,
            case_sensitive: None,
            column_sort: None,
            r#ref: None,
            sort_method: None,
        };
        let mut conditions: Vec<XlsxSortCondition> = vec![];

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"caseSensitive" => {
                            sort_state.case_sensitive = string_to_bool(&string_value);
                        }
                        b"columnSort" => {
                            sort_state.column_sort = string_to_bool(&string_value);
                        }
                        b"ref" => {
                            let value = a.value.as_ref();
                            sort_state.r#ref = Dimension::from_a1(value);
                        }
                        b"sortMethod" => {
                            sort_state.sort_method = Some(string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sortCondition" => {
                    conditions.push(XlsxSortCondition::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"sortState" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `sortState`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        sort_state.sort_condition = Some(conditions);

        Ok(sort_state)
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sortcondition?view=openxml-3.0.1
///
/// Sort condition.
/// When more than one sortCondition is specified, the first condition is applied first, then the second condition is applied, and so on.
///
/// Example:
/// ```
/// <sortCondition descending="1" ref="B4:B8"/>
/// ```
///
/// sortCondition (Sort Condition)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxSortCondition {
    // Attributes	Description
    /// customList (Custom List)	Sort by a custom list.
    pub custom_list: Option<String>,

    /// descending (Descending)	Sort descending.
    pub descending: Option<bool>,

    /// dxfId (Format Id)
    ///
    /// Format Id when sortBy=cellColor or fontColor
    pub dxf_id: Option<u64>,

    // The possible values for this attribute are defined by the ST_DxfId simple type (§18.18.25).
    /// iconId (Icon Id)
    ///
    /// Zero-based index of an icon in an icon set.
    /// The absence of this attribute means "no icon"
    pub icon_id: Option<u64>,

    /// iconSet (Icon Set)
    ///
    /// Icon set index when sortBy=icon.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.iconsetvalues?view=openxml-3.0.1
    pub icon_set: Option<String>,

    /// ref (Reference)
    ///
    /// Column/Row that this sort condition applies to.
    /// This shall be contained within the ref in CT_SortState.
    pub r#ref: Option<Dimension>,

    /// sortBy (Sort By)
    ///
    /// Type of sort.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sortbyvalues?view=openxml-3.0.1
    pub sort_by: Option<String>,
}

impl XlsxSortCondition {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut condition = Self {
            custom_list: None,
            descending: None,
            dxf_id: None,
            icon_id: None,
            icon_set: None,
            r#ref: None,
            sort_by: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"customList" => condition.custom_list = Some(string_value),
                        b"descending" => {
                            condition.descending = string_to_bool(&string_value);
                        }
                        b"dxfId" => {
                            condition.dxf_id = string_to_unsignedint(&string_value);
                        }
                        b"iconId" => {
                            condition.icon_id = string_to_unsignedint(&string_value);
                        }
                        b"iconSet" => {
                            condition.icon_set = Some(string_value);
                        }
                        b"ref" => {
                            let value = a.value.as_ref();
                            condition.r#ref = Dimension::from_a1(value);
                        }
                        b"sortBy" => {
                            condition.sort_by = Some(string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        Ok(condition)
    }
}
