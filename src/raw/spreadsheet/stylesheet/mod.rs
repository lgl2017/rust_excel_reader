pub mod border;
pub mod cell_style;
pub mod color;
pub mod fill;
pub mod font;
pub mod format;
pub mod table_style;

use std::io::{Read, Seek};

use anyhow::bail;
use quick_xml::events::Event;
use zip::ZipArchive;

use crate::excel::xml_reader;

use border::{load_borders, Border, Borders};
use cell_style::{load_cell_styles, CellStyles};
use color::stylesheet_colors::StyleSheetColors;
use fill::{load_fills, Fill, Fills};
use font::{load_fonts, Font, Fonts};
use format::{
    cell_format::CellFormat,
    cell_style_xfs::{load_cell_styles_xfs, CellStyleFormats},
    cell_xfs::{load_cell_xfs, CellFormats},
    differential_format::{load_dxfs, DifferentialFormats},
    numbering_format::{load_number_formats, NumberingFormat, NumberingFormats},
};
use table_style::TableStyles;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.stylesheet?view=openxml-3.0.1
///
/// Root element of the StylesSheet part
///
/// tag: styleSheet
#[derive(Debug, Clone, PartialEq)]
pub struct StyleSheet {
    // extLst (Future Feature Data Storage Area): Not supported

    // children
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fills?view=openxml-3.0.1
    pub fills: Option<Fills>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.borders?view=openxml-3.0.1
    pub borders: Option<Borders>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.colors?view=openxml-3.0.1
    pub colors: Option<StyleSheetColors>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fonts?view=openxml-3.0.1
    pub fonts: Option<Fonts>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyles?view=openxml-3.0.1
    ///
    /// This element represents the name and related formatting records for a named cell style in this workbook.
    // tag: cellStyles
    pub cell_styles: Option<CellStyles>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyleformats?view=openxml-3.0.1
    ///
    /// This element contains the master formatting records (xf's) which define the formatting for all named cell styles in this workbook.
    /// Master formatting records reference individual elements of formatting (e.g., number format, font definitions, cell fills, etc.) by specifying a zero-based index into those collections.
    /// Master formatting records also specify whether to apply or ignore particular aspects of formatting.
    // tag: cellStyleXfs
    pub cell_style_xfs: Option<CellStyleFormats>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformats?view=openxml-3.0.1
    ///
    /// This element contains the master formatting records (xf) which define the formatting applied to cells in this workbook.
    /// These records are the starting point for determining the formatting for a cell.
    /// Cells in the Sheet Part reference the xf records by zero-based index.
    // tag: cellXfs
    pub cell_xfs: Option<CellFormats>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.differentialformats?view=openxml-3.0.1
    ///
    /// This element contains the master differential formatting records (dxf's) which define formatting for all non-cell formatting in this workbook.
    /// tag: dxfs
    pub differential_xfs: Option<DifferentialFormats>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformats?view=openxml-3.0.1
    ///
    /// This element defines the number formats in this workbook, consisting of a sequence of numFmt records,
    /// where each numFmt record defines a particular number format, indicating how to format and render the numeric value of a cell.
    // tag: numFmts
    pub numbering_formats: Option<NumberingFormats>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyles?view=openxml-3.0.1
    // tableStyles
    pub table_styles: Option<TableStyles>,
}

impl StyleSheet {
    pub(crate) fn load(zip: &mut ZipArchive<impl Read + Seek>) -> anyhow::Result<Self> {
        let path = "xl/styles.xml";
        let mut style_sheet = Self {
            fills: None,
            borders: None,
            colors: None,
            fonts: None,
            cell_styles: None,
            cell_style_xfs: None,
            cell_xfs: None,
            differential_xfs: None,
            numbering_formats: None,
            table_styles: None,
        };

        let Some(mut reader) = xml_reader(zip, path) else {
            return Ok(style_sheet);
        };

        let mut buf = Vec::new();

        loop {
            buf.clear();

            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"extLst" => {
                    let _ = reader.read_to_end_into(e.to_end().to_owned().name(), &mut Vec::new());
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fills" => {
                    let fills = load_fills(&mut reader)?;
                    style_sheet.fills = Some(fills);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"borders" => {
                    let borders = load_borders(&mut reader)?;
                    style_sheet.borders = Some(borders);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"colors" => {
                    let colors = StyleSheetColors::load(&mut reader)?;
                    style_sheet.colors = Some(colors);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"fonts" => {
                    let fonts = load_fonts(&mut reader)?;
                    style_sheet.fonts = Some(fonts);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cellStyles" => {
                    let styles = load_cell_styles(&mut reader)?;
                    style_sheet.cell_styles = Some(styles);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cellStyleXfs" => {
                    let formats = load_cell_styles_xfs(&mut reader)?;
                    style_sheet.cell_style_xfs = Some(formats);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"cellXfs" => {
                    let formats = load_cell_xfs(&mut reader)?;
                    style_sheet.cell_xfs = Some(formats);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"dxfs" => {
                    let formats = load_dxfs(&mut reader)?;
                    style_sheet.differential_xfs = Some(formats);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"numFmts" => {
                    let formats = load_number_formats(&mut reader)?;
                    style_sheet.numbering_formats = Some(formats);
                }
                Ok(Event::Start(ref e)) if e.local_name().as_ref() == b"tableStyles" => {
                    let styles = TableStyles::load(&mut reader, e)?;
                    style_sheet.table_styles = Some(styles);
                }
                Ok(Event::End(ref e)) if e.local_name().as_ref() == b"styleSheet" => break,
                Ok(Event::Eof) => break,
                Err(e) => bail!(e.to_string()),
                _ => (),
            }
        }

        return Ok(style_sheet);
    }
}

impl StyleSheet {
    /// Get cell format (cell_xfs) by a given style index (0 based) specified in cell / col / row.
    pub(crate) fn get_cell_format(&self, xf_id: usize) -> Option<CellFormat> {
        let cell_xfs = self.cell_xfs.clone().unwrap_or(vec![]);
        if xf_id >= cell_xfs.len() {
            return None;
        }
        return Some(cell_xfs[xf_id].clone());
    }
    /// Get cell style format (cellStyleXfs) by a given style index (0 based)
    pub(crate) fn get_cell_style_format(&self, xf_id: usize) -> Option<CellFormat> {
        let cell_style_xfs = self.cell_style_xfs.clone().unwrap_or(vec![]);
        if xf_id >= cell_style_xfs.len() {
            return None;
        }
        return Some(cell_style_xfs[xf_id].clone());
    }

    /// Get font by a given font index (0 based).
    pub(crate) fn get_font(&self, index: usize) -> Option<Font> {
        let fonts = self.fonts.clone().unwrap_or(vec![]);
        if index >= fonts.len() {
            return None;
        }
        return Some(fonts[index].clone());
    }

    /// Get border by a given border index (0 based).
    pub(crate) fn get_border(&self, index: usize) -> Option<Border> {
        let borders = self.borders.clone().unwrap_or(vec![]);
        if index >= borders.len() {
            return None;
        }
        return Some(borders[index].clone());
    }

    /// Get fill by a given fill index (0 based).
    pub(crate) fn get_fill(&self, index: usize) -> Option<Fill> {
        let fills = self.fills.clone().unwrap_or(vec![]);
        if index >= fills.len() {
            return None;
        }
        return Some(fills[index].clone());
    }

    /// Get numbering format code by a given number_format_id
    pub(crate) fn get_num_format(&self, num_format_id: u64) -> Option<NumberingFormat> {
        let numbering_formats = self.numbering_formats.clone().unwrap_or(vec![]);

        let filtered: Vec<NumberingFormat> = numbering_formats
            .into_iter()
            .filter(|n| n.num_fmt_id == Some(num_format_id))
            .collect();

        return filtered.first().cloned();
    }
}
