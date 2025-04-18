use anyhow::bail;
use quick_xml::events::{BytesStart, Event};

use super::{filter_column::XlsxFilterColumn, sort_state::XlsxSortState};
use crate::{common_types::Dimension, excel::XmlReader};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.autofilter?view=openxml-3.0.1
///
/// AutoFilter temporarily hides rows based on a filter criteria, which is applied column by column to a table of data in the worksheet.
/// This collection expresses AutoFilter settings.
///
/// Example:
/// ```
/// <autoFilter ref="B3:E8">
///     <filterColumn colId="0">
///         <customFilters>
///             <customFilter operator="greaterThan" val="0.5"/>
///         </customFilters>
///     </filterColumn>
///     <filterColumn colId="1" hiddenButton="1" />
/// </autoFilter>
/// ```
///
/// autoFilter (AutoFilter Settings)
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxAutoFilter {
    // extLst (Future Feature Data Storage Area) not supported

    // Child Elements
    /// filterColumn (AutoFilter Column)
    pub filter_colomn: Option<Vec<XlsxFilterColumn>>,

    /// sortState (Sort State)
    pub sort_state: Option<XlsxSortState>,

    /// Attributes
    /// ref (Cell or Range Reference)
    ///
    /// Reference to the cell range to which the AutoFilter is applied.
    pub r#ref: Option<Dimension>,
}

impl XlsxAutoFilter {
    pub(crate) fn load(reader: &mut XmlReader, e: &BytesStart) -> anyhow::Result<Self> {
        let attributes = e.attributes();
        let mut filter = Self {
            filter_colomn: None,
            sort_state: None,
            r#ref: None,
        };
        let mut filter_columns: Vec<XlsxFilterColumn> = vec![];

        for a in attributes {
            match a {
                Ok(a) => match a.key.local_name().as_ref() {
                    b"ref" => {
                        let value = a.value.as_ref();
                        filter.r#ref = Dimension::from_a1(value);
                        break;
                    }
                    _ => {}
                },
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
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"filterColumn" => {
                    filter_columns.push(XlsxFilterColumn::load(reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sortState" => {
                    filter.sort_state = Some(XlsxSortState::load(reader, e)?);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"autoFilter" => break,
                Ok(Event::Eof) => bail!("unexpected end of file at `autoFilter`."),
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        filter.filter_colomn = Some(filter_columns);

        Ok(filter)
    }
}
