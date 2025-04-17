use std::io::{Read, Seek};

use anyhow::bail;
use column_information::{load_column_infos, ColumnInformations};
use merge_cell::{load_merge_cells, MergeCells};
use quick_xml::events::Event;
use sheet_data::SheetData;
use sheet_dimension::{load_sheet_dimension, SheetDimension};
use table_part::{load_table_parts, TableParts};
use zip::ZipArchive;

use crate::{
    excel::xml_reader,
    raw::spreadsheet::{
        filter::auto_filter::AutoFilter, string_item::phonetic_properties::PhoneticProperties,
    },
};

use super::sheet_format_properties::SheetFormatProperties;

pub mod cell;
pub mod column_information;
pub mod merge_cell;
pub mod row;
pub mod sheet_data;
pub mod sheet_dimension;
pub mod sheet_view;
pub mod table_part;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.worksheet?view=openxml-3.0.1
///
/// This is the root element of Worksheet parts.
/// Contains information on dimension, sheetview, column styles, cell data, cell styles, merged regions, etc.
///
/// Example:
/// ```
/// <worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
///   <dimension ref="A2:B7" />
///   <sheetViews>
///     <sheetView workbookViewId="0" showGridLines="0" defaultGridColor="1">
///       <pane topLeftCell="B3" xSplit="1" ySplit="2" activePane="bottomRight" state="frozen" />
///     </sheetView>
///   </sheetViews>
///   <sheetFormatPr defaultColWidth="16.3333" defaultRowHeight="19.9" customHeight="1" outlineLevelRow="0" outlineLevelCol="0" />
///   <cols>
///     <col min="1" max="2" width="16.3516" style="1" customWidth="1" />
///     <col min="3" max="16384" width="16.3516" style="1" customWidth="1" />
///   </cols>
///   <sheetData>
///     <row r="1" ht="27.65" customHeight="1">
///       <c r="A1" t="s" s="2">
///         <v>0</v>
///       </c>
///       <c r="B1" s="2" />
///     </row>
///     /// ...
///   </sheetData>
///   <mergeCells count="2">
///     <mergeCell ref="A1:B1" />
///     <mergeCell ref="B6:B7" />
///   </mergeCells>
///   <pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.277778" footer="0.277778" />
///   <pageSetup firstPageNumber="1" fitToHeight="1" fitToWidth="1" scale="72" useFirstPageNumber="0" orientation="portrait" pageOrder="downThenOver" />
///   <headerFooter>
///     <oddFooter>&amp;C&amp;"Helvetica Neue,Regular"&amp;12&amp;K000000&amp;P</oddFooter>
///   </headerFooter>
/// </worksheet>
/// ```
/// worksheet (Worksheet)
#[derive(Debug, Clone, PartialEq)]
pub struct Worksheet {
    // extLst (Future Feature Data Storage Area) Not supported

    // Child Elements	Subclause
    // autoFilter (AutoFilter Settings)	§18.3.1.2
    pub auto_filter: Option<AutoFilter>,
    // cellWatches (Cell Watch Items)	§18.3.1.9
    // colBreaks (Vertical Page Breaks)	§18.3.1.14

    // cols (Column Information)	§18.3.1.17
    pub column_infos: Option<ColumnInformations>,
    // conditionalFormatting (Conditional Formatting)	§18.3.1.18
    // controls (Embedded Controls)	§18.3.1.21
    // customProperties (Custom Properties)	§18.3.1.23
    // customSheetViews (Custom Sheet Views)	§18.3.1.27
    // dataConsolidate (Data Consolidate)	§18.3.1.29
    // dataValidations (Data Validations)	§18.3.1.33

    // dimension (Worksheet Dimensions)	§18.3.1.35
    pub dimension: Option<SheetDimension>,
    // drawing (Drawing)	§18.3.1.36
    // drawingHF (Drawing Reference in Header Footer)	§18.3.1.37
    // headerFooter (Header Footer Settings)	§18.3.1.46
    // hyperlinks (Hyperlinks)	§18.3.1.48
    // ignoredErrors (Ignored Errors)	§18.3.1.51

    // mergeCells (Merge Cells)	§18.3.1.55
    pub merge_cells: Option<MergeCells>,

    // oleObjects (Embedded Objects)	§18.3.1.60
    // pageMargins (Page Margins)	§18.3.1.62
    // pageSetup (Page Setup Settings)	§18.3.1.63
    // phoneticPr (Phonetic Properties)	§18.4.3
    pub phonetic_properties: Option<PhoneticProperties>,

    // picture (Background Image)	§18.3.1.67
    // printOptions (Print Options)	§18.3.1.70
    // protectedRanges (Protected Ranges)	§18.3.1.72
    // rowBreaks (Horizontal Page Breaks (Row))	§18.3.1.74
    // scenarios (Scenarios)	§18.3.1.76
    // sheetCalcPr (Sheet Calculation Properties)	§18.3.1.79

    // sheetData (Sheet Data)	§18.3.1.80
    pub sheet_data: Option<SheetData>,

    // sheetFormatPr (Sheet Format Properties)	§18.3.1.81
    pub sheet_format_properties: Option<SheetFormatProperties>,
    // sheetPr (Sheet Properties)	§18.3.1.82
    // sheetProtection (Sheet Protection Options)	§18.3.1.85
    // sheetViews (Sheet Views)	§18.3.1.88
    // smartTags (Smart Tags)	§18.3.1.90
    // sortState (Sort State)	§18.3.1.92

    // tableParts (Table Parts)	§18.3.1.95
    pub table_parts: Option<TableParts>, // webPublishItems (Web Publishing Items)
}

impl Worksheet {
    pub(crate) fn load(zip: &mut ZipArchive<impl Read + Seek>, path: &str) -> anyhow::Result<Self> {
        let mut worksheet = Self {
            auto_filter: None,
            column_infos: None,
            dimension: None,
            merge_cells: None,
            phonetic_properties: None,
            sheet_data: None,
            sheet_format_properties: None,
            table_parts: None,
        };

        let Some(mut reader) = xml_reader(zip, path) else {
            return Ok(worksheet);
        };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"autoFilter" => {
                    worksheet.auto_filter = Some(AutoFilter::load(&mut reader, e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cols" => {
                    worksheet.column_infos = Some(load_column_infos(&mut reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dimension" => {
                    worksheet.dimension = load_sheet_dimension(e)?;
                }

                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"mergeCells" => {
                    worksheet.merge_cells = Some(load_merge_cells(&mut reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"phoneticPr" => {
                    worksheet.phonetic_properties = Some(PhoneticProperties::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheetData" => {
                    worksheet.sheet_data = Some(SheetData::load(&mut reader)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"sheetFormatPr" => {
                    worksheet.sheet_format_properties = Some(SheetFormatProperties::load(e)?);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tableParts" => {
                    worksheet.table_parts = Some(load_table_parts(&mut reader)?);
                }

                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"worksheet" => break,
                Ok(Event::Eof) => break,
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        return Ok(worksheet);
    }
}
