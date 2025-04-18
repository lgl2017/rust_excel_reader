use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use crate::{
    excel::XmlReader,
    helper::{extract_val_attribute, string_to_bool, string_to_unsignedint},
};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.filters?view=openxml-3.0.1
///
/// When multiple values are chosen to filter by, or when a group of date values are chosen to filter by, this element groups those criteria together.
///
/// Example
/// ```
/// <filters>
///   <dateGroupItem year="2006" month="1" day="2" dateTimeGrouping="day"/>
///   <dateGroupItem year="2005" month="1" day="2" dateTimeGrouping="day"/>
/// </filters>
///
/// <filters>
///   <filter val="0.316588716"/>
///   <filter val="0.667439395"/>
///   <filter val="0.823086999"/>
/// </filters>
/// ```
///
/// filters (Filter Criteria)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxFilterCriteriaGroup {
    // Child Elements
    /// dateGroupItem (Date Grouping)
    pub date_group_item: Option<Vec<XlsxDateGroupItem>>,
    /// filter (Filter)	§18.3.2.6
    pub value_filters: Option<Vec<XlsxValueFilter>>,

    // Attributes
    /// blank (Filter by Blank)
    ///
    /// Flag indicating whether to filter by blank.
    pub filter_by_blank: Option<bool>,

    /// calendarType (Calendar Type)
    ///
    /// Calendar type for date grouped items. Used to interpret the values in dateGroupItem.
    /// This is the calendar type used to evaluate all dates in the filter column, even when those dates are not using the same calendar system / date formatting.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.calendarvalues?view=openxml-3.0.1
    pub calendar_type: Option<String>,
}

impl XlsxFilterCriteriaGroup {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let mut creteria = Self {
            date_group_item: None,
            value_filters: None,
            filter_by_blank: None,
            calendar_type: None,
        };

        let mut date_group_items: Vec<XlsxDateGroupItem> = vec![];
        let mut value_filters: Vec<XlsxValueFilter> = vec![];

        let attributes = e.attributes();
        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"blank" => {
                            creteria.filter_by_blank = string_to_bool(&string_value);
                        }
                        b"calendarType" => {
                            creteria.calendar_type = Some(string_value);
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"filter" => {
                    value_filters.push(XlsxValueFilter::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dateGroupItem" => {
                    date_group_items.push(XlsxDateGroupItem::load(e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"filters" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `filters`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        creteria.value_filters = Some(value_filters);
        creteria.date_group_item = Some(date_group_items);

        return Ok(creteria);
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.filter?view=openxml-3.0.1
///
/// This element expresses a filter criteria value.
///
/// Example
/// ```
/// <filter val="0.823086999"/>
/// ```
///
/// filter (Filter)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxValueFilter {
    // Attributes
    /// val (Filter Value)
    ///
    /// Filter value used in the criteria.
    pub filter_value: Option<String>,
}

impl XlsxValueFilter {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let val = extract_val_attribute(e)?;
        Ok(Self { filter_value: val })
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.dategroupitem?view=openxml-3.0.1
///
/// This collection is used to express a group of dates or times which are used in an AutoFilter criteria.
/// Values are always written in the calendar type of the first date encountered in the filter range, so that all subsequent dates, even when formatted or represented by other calendar types, can be correctly compared for the purposes of filtering.
///
/// Example:
/// ```
/// <dateGroupItem year="2006" month="1" day="2" dateTimeGrouping="day"/>
/// ```
///
/// dateGroupItem (Date Grouping)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxDateGroupItem {
    // Attributes
    /// dateTimeGrouping (Date Time Grouping)
    ///
    /// Grouping level.
    ///
    /// possible values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.datetimegroupingvalues?view=openxml-3.0.1
    pub grouping_level: Option<String>,

    /// day (Day)
    ///
    /// Day (1-31)
    pub day: Option<u64>,

    /// hour (Hour)
    ///
    /// Hour (0-23)
    pub hour: Option<u64>,

    /// minute (Minute)
    ///
    /// Minute (0-59)
    pub minute: Option<u64>,

    /// month (Month)
    ///
    /// Month (1-12)
    pub month: Option<u64>,

    /// second (Second)
    ///
    /// Second (0-59)
    pub second: Option<u64>,

    /// year (Year)
    ///
    /// Year (4 digits)
    pub year: Option<u64>,
}

impl XlsxDateGroupItem {
    pub(crate) fn load(e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut item = Self {
            grouping_level: None,
            day: None,
            hour: None,
            minute: None,
            month: None,
            second: None,
            year: None,
        };

        for a in attributes {
            match a {
                Ok(a) => {
                    let string_value = String::from_utf8(a.value.to_vec())?;
                    match a.key.local_name().as_ref() {
                        b"dateTimeGrouping" => item.grouping_level = Some(string_value),
                        b"day" => {
                            item.day = string_to_unsignedint(&string_value);
                        }
                        b"hour" => {
                            item.hour = string_to_unsignedint(&string_value);
                        }
                        b"minute" => {
                            item.minute = string_to_unsignedint(&string_value);
                        }
                        b"month" => {
                            item.month = string_to_unsignedint(&string_value);
                        }
                        b"second" => {
                            item.second = string_to_unsignedint(&string_value);
                        }
                        b"year" => {
                            item.year = string_to_unsignedint(&string_value);
                        }
                        _ => {}
                    }
                }
                Err(error) => {
                    bail!(error.to_string())
                }
            }
        }
        Ok(item)
    }
}
