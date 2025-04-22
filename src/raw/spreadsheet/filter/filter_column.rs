use std::io::Read;
use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    helper::{string_to_bool, string_to_unsignedint},
};

use super::{
    color_filter::XlsxColorFilter, custom_filter::XlsxCustomFilters,
    dynamic_filter::XlsxDynamicFilter, filter_criteria::XlsxFilterCriteriaGroup,
    icon_filter::XlsxIconFilter, top_n_filter::XlsxTopNFilter,
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.filtercolumn?view=openxml-3.0.1
///
/// The filterColumn collection identifies a particular column in the AutoFilter range and specifies filter information that has been applied to this column.
/// If a column in the AutoFilter range has no criteria specified, then there is no corresponding filterColumn collection expressed for that column.
///
/// Exampl
/// ```
/// <filterColumn colId="0">
///     <customFilters and="1">
///         <customFilter operator="greaterThanOrEqual" val="0.2"/>
///         <customFilter operator="lessThanOrEqual" val="0.5"/>
///     </customFilters>
///     <dynamicFilter type="today"/>
///     <colorFilter dxfId="0" cellColor="0"/>
///     <filters>
///         <dateGroupItem year="2006" month="1" day="2" dateTimeGrouping="day"/>
///         <dateGroupItem year="2005" month="1" day="2" dateTimeGrouping="day"/>
///     </filters>
///     <iconFilter iconSet="3Arrows" iconId="0"/>
/// </filterColumn>
/// ```
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxFilterColumn {
    // extLst (Future Feature Data Storage Area) Not supported

    // Child Elements
    /// colorFilter (Color Filter Criteria)
    pub color_filter: Option<XlsxColorFilter>,

    /// customFilters (Custom Filters)
    pub custom_filters: Option<XlsxCustomFilters>,

    /// dynamicFilter (Dynamic Filter)
    pub dynamic_filter: Option<XlsxDynamicFilter>,

    /// filters (Filter Criteria)
    pub grouped_filter: Option<XlsxFilterCriteriaGroup>,

    /// iconFilter (Icon Filter)
    pub icon_filter: Option<XlsxIconFilter>,

    /// top10 (Top 10)
    pub top_n_filter: Option<XlsxTopNFilter>,

    // Attributes
    /// colId (Filter Column Data)
    ///
    /// Zero-based index indicating the AutoFilter column to which this filter information applies.
    pub col_id: Option<u64>,

    /// hiddenButton (Hidden AutoFilter Button)
    ///
    /// Flag indicating whether the AutoFilter button for this column is hidden.
    pub hidden_autofilter_button: Option<bool>,

    /// showButton (Show Filter Button)
    ///
    /// Flag indicating whether the filter button is visible.
    pub show_filter_button: Option<bool>,
}

impl XlsxFilterColumn {
    pub(crate) fn load(reader: &mut XmlReader<impl Read>, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut filter_column = Self {
            color_filter: None,
            custom_filters: None,
            dynamic_filter: None,
            grouped_filter: None,
            icon_filter: None,
            top_n_filter: None,
            col_id: None,
            hidden_autofilter_button: None,
            show_filter_button: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"colId" => {
                            filter_column.col_id = string_to_unsignedint(&string_value);
                        }
                        b"hiddenButton" => {
                            filter_column.hidden_autofilter_button = string_to_bool(&string_value);
                        }
                        b"showButton" => {
                            filter_column.show_filter_button = string_to_bool(&string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"colorFilter" => {
                    filter_column.color_filter = Some(XlsxColorFilter::load(e)?)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"customFilters" => {
                    filter_column.custom_filters = Some(XlsxCustomFilters::load(reader, e)?)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dynamicFilter" => {
                    filter_column.dynamic_filter = Some(XlsxDynamicFilter::load(e)?)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"filters" => {
                    filter_column.grouped_filter = Some(XlsxFilterCriteriaGroup::load(reader, e)?)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"iconFilter" => {
                    filter_column.icon_filter = Some(XlsxIconFilter::load(e)?)
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"top10" => {
                    filter_column.top_n_filter = Some(XlsxTopNFilter::load(e)?)
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"filterColumn" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `filterColumn`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        Ok(filter_column)
    }
}
