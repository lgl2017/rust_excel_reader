pub mod border;
pub mod cell_style;
pub mod color;
pub mod fill;
pub mod font;
pub mod format;
pub mod table_style;

use anyhow::bail;
use quick_xml::events::Event;
use std::io::{Read, Seek};
use zip::ZipArchive;

use crate::excel::xml_reader;

use border::{load_borders, XlsxBorder, XlsxBorders};
use cell_style::{load_cell_styles, XlsxCellStyles};
use color::stylesheet_colors::XlsxStyleSheetColors;
use fill::{load_fills, XlsxFill, XlsxFills};
use font::{load_fonts, XlsxFont, XlsxFonts};
use format::{
    cell_format::XlsxCellFormat,
    cell_style_xfs::{load_cell_styles_xfs, XlsxCellStyleFormats},
    cell_xfs::{load_cell_xfs, XlsxCellFormats},
    differential_format::{load_dxfs, XlsxDifferentialFormats},
    numbering_format::{load_number_formats, XlsxNumberingFormat, XlsxNumberingFormats},
};
use table_style::XlsxTableStyles;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.stylesheet?view=openxml-3.0.1
///
/// Root element of the StylesSheet part
///
/// tag: styleSheet
#[derive(Debug, Clone, PartialEq)]
pub struct XlsxStyleSheet {
    // extLst (Future Feature Data Storage Area): Not supported

    // children
    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fills?view=openxml-3.0.1
    pub fills: Option<XlsxFills>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.borders?view=openxml-3.0.1
    pub borders: Option<XlsxBorders>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.colors?view=openxml-3.0.1
    pub colors: Option<XlsxStyleSheetColors>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.fonts?view=openxml-3.0.1
    pub fonts: Option<XlsxFonts>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyles?view=openxml-3.0.1
    ///
    /// This element represents the name and related formatting records for a named cell style in this workbook.
    // tag: cellStyles
    pub cell_styles: Option<XlsxCellStyles>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellstyleformats?view=openxml-3.0.1
    ///
    /// This element contains the master formatting records (xf's) which define the formatting for all named cell styles in this workbook.
    /// Master formatting records reference individual elements of formatting (e.g., number format, font definitions, cell fills, etc.) by specifying a zero-based index into those collections.
    /// Master formatting records also specify whether to apply or ignore particular aspects of formatting.
    // tag: cellStyleXfs
    pub cell_style_xfs: Option<XlsxCellStyleFormats>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellformats?view=openxml-3.0.1
    ///
    /// This element contains the master formatting records (xf) which define the formatting applied to cells in this workbook.
    /// These records are the starting point for determining the formatting for a cell.
    /// Cells in the Sheet Part reference the xf records by zero-based index.
    // tag: cellXfs
    pub cell_xfs: Option<XlsxCellFormats>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.differentialformats?view=openxml-3.0.1
    ///
    /// This element contains the master differential formatting records (dxf's) which define formatting for all non-cell formatting in this workbook.
    /// tag: dxfs
    pub differential_xfs: Option<XlsxDifferentialFormats>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformats?view=openxml-3.0.1
    ///
    /// This element defines the number formats in this workbook, consisting of a sequence of numFmt records,
    /// where each numFmt record defines a particular number format, indicating how to format and render the numeric value of a cell.
    // tag: numFmts
    pub numbering_formats: Option<XlsxNumberingFormats>,

    /// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.tablestyles?view=openxml-3.0.1
    // tableStyles
    pub table_styles: Option<XlsxTableStyles>,
}

impl XlsxStyleSheet {
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
                    let colors = XlsxStyleSheetColors::load(&mut reader)?;
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
                    let styles = XlsxTableStyles::load(&mut reader, e)?;
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

impl XlsxStyleSheet {
    /// Get cell format (cell_xfs) by a given style index (0 based) specified in cell / col / row.
    pub(crate) fn get_cell_format(&self, xf_id: usize) -> Option<XlsxCellFormat> {
        let cell_xfs = self.cell_xfs.clone().unwrap_or(vec![]);
        if xf_id >= cell_xfs.len() {
            return None;
        }
        return Some(cell_xfs[xf_id].clone());
    }
    /// Get cell style format (cellStyleXfs) by a given style index (0 based)
    pub(crate) fn get_cell_style_format(&self, xf_id: usize) -> Option<XlsxCellFormat> {
        let cell_style_xfs = self.cell_style_xfs.clone().unwrap_or(vec![]);
        if xf_id >= cell_style_xfs.len() {
            return None;
        }
        return Some(cell_style_xfs[xf_id].clone());
    }

    /// Get font by a given font index (0 based).
    pub(crate) fn get_font(&self, index: usize) -> Option<XlsxFont> {
        let fonts = self.fonts.clone().unwrap_or(vec![]);
        if index >= fonts.len() {
            return None;
        }
        return Some(fonts[index].clone());
    }

    /// Get border by a given border index (0 based).
    pub(crate) fn get_border(&self, index: usize) -> Option<XlsxBorder> {
        let borders = self.borders.clone().unwrap_or(vec![]);
        if index >= borders.len() {
            return None;
        }
        return Some(borders[index].clone());
    }

    /// Get fill by a given fill index (0 based).
    pub(crate) fn get_fill(&self, index: usize) -> Option<XlsxFill> {
        let fills = self.fills.clone().unwrap_or(vec![]);
        if index >= fills.len() {
            return None;
        }
        return Some(fills[index].clone());
    }

    /// Get numbering format code by a given number_format_id
    pub(crate) fn get_num_format(&self, num_format_id: u64) -> Option<XlsxNumberingFormat> {
        let numbering_formats = self.numbering_formats.clone().unwrap_or(vec![]);

        let filtered: Vec<XlsxNumberingFormat> = numbering_formats
            .into_iter()
            .filter(|n| n.num_fmt_id == Some(num_format_id))
            .collect();

        return filtered.first().cloned();
    }
}
