use border::Border;
use fill::Fill;
use font::Font;
use numbering_format::NumberingFormat;
use text_alignment::TextAlignment;

use crate::raw::{
    drawing::scheme::color_scheme::ColorScheme,
    spreadsheet::{
        sheet::{
            sheet_format_properties::SheetFormatProperties as RawSheetProperties,
            worksheet::{
                cell::Cell as RawCell, column_information::ColumnInformation as RawCol,
                row::Row as RawRow,
            },
        },
        stylesheet::{
            format::{alignment::Alignment, protection::Protection},
            StyleSheet,
        },
    },
};

pub mod border;
pub mod fill;
pub mod font;
pub mod numbering_format;
pub mod text_alignment;

static DEFAULT_CELL_WIDTH: f64 = 8.43;
static DEFAULT_CELL_HEIGHT: f64 = 15.0;

#[derive(Debug, Clone, PartialEq)]
pub struct CellProperty {
    pub width: f64,
    /// Flag indicating if the column width should automatically resize
    pub width_best_fit: bool,

    pub height: f64,
    pub hidden: bool,
    pub show_phonetic: bool,
    pub alignment: TextAlignment,
    pub font: Font,
    pub border: Border,
    pub fill: Fill,
    pub numbering_format: NumberingFormat,
}

impl CellProperty {
    pub(crate) fn from_raw(
        cell: RawCell,
        row_info: RawRow,
        col_info: Option<RawCol>,
        fill_id: Option<u64>,
        font_id: Option<u64>,
        border_id: Option<u64>,
        numbering_format_id: Option<u64>,
        alignment: Option<Alignment>,
        protection: Option<Protection>,
        sheet_format_properties: Option<RawSheetProperties>,
        stylesheet: StyleSheet,
        color_scheme: Option<ColorScheme>,
    ) -> Self {
        let show_phonetic = Self::show_phonetic(cell.clone(), col_info.clone(), row_info.clone());
        let width = Self::cell_width(col_info.clone(), sheet_format_properties.clone());
        let height = Self::cell_height(row_info.clone(), sheet_format_properties.clone());

        let hidden = Self::cell_hidden(
            protection.clone(),
            col_info.clone(),
            row_info.clone(),
            sheet_format_properties.clone(),
        );

        let fill = Self::get_fill(fill_id, stylesheet.clone(), color_scheme.clone());
        let font = Self::get_font(font_id, stylesheet.clone(), color_scheme.clone());
        let border = Self::get_border(border_id, stylesheet.clone(), color_scheme.clone());
        let numbering_format = Self::get_numbering_format(numbering_format_id, stylesheet.clone());
        return Self {
            width,
            width_best_fit: Self::get_width_best_fit(col_info),
            height,
            hidden,
            show_phonetic,
            alignment: TextAlignment::from_raw(alignment),
            font,
            border,
            fill,
            numbering_format,
        };
    }

    fn get_width_best_fit(col_info: Option<RawCol>) -> bool {
        let Some(col_info) = col_info else {
            return false;
        };

        return col_info.best_fit.unwrap_or(false);
    }

    fn get_numbering_format(
        numbering_format: Option<u64>,
        stylesheet: StyleSheet,
    ) -> NumberingFormat {
        if let Some(id) = numbering_format {
            let raw = stylesheet.get_num_format(id);
            return NumberingFormat::from_raw(raw);
        }

        return NumberingFormat::default();
    }

    fn get_font(
        font_id: Option<u64>,
        stylesheet: StyleSheet,
        color_scheme: Option<ColorScheme>,
    ) -> Font {
        if let Some(id) = font_id {
            if let Ok(id) = TryInto::<usize>::try_into(id) {
                let raw = stylesheet.get_font(id);
                return Font::from_raw_font(raw, stylesheet.colors.clone(), color_scheme);
            }
        }

        return Font::default();
    }

    fn get_border(
        border_id: Option<u64>,
        stylesheet: StyleSheet,
        color_scheme: Option<ColorScheme>,
    ) -> Border {
        if let Some(id) = border_id {
            if let Ok(id) = TryInto::<usize>::try_into(id) {
                let raw = stylesheet.get_border(id);
                return Border::from_raw(raw, stylesheet.colors.clone(), color_scheme);
            }
        }

        return Border::default();
    }

    fn get_fill(
        fill_id: Option<u64>,
        stylesheet: StyleSheet,
        color_scheme: Option<ColorScheme>,
    ) -> Fill {
        if let Some(id) = fill_id {
            if let Ok(id) = TryInto::<usize>::try_into(id) {
                let raw = stylesheet.get_fill(id);
                return Fill::from_raw(raw, stylesheet.colors.clone(), color_scheme);
            }
        }

        return Fill::default();
    }

    fn show_phonetic(cell: RawCell, col_info: Option<RawCol>, row_info: RawRow) -> bool {
        return if let Some(b) = cell.show_phonetic {
            b
        } else if let Some(b) = row_info.show_phonetic {
            b
        } else if let Some(col_info) = col_info {
            if let Some(b) = col_info.show_phonetic {
                b
            } else {
                true
            }
        } else {
            true
        };
    }

    fn cell_width(
        col_info: Option<RawCol>,
        sheet_format_properties: Option<RawSheetProperties>,
    ) -> f64 {
        if let Some(col_info) = col_info {
            if let Some(f) = col_info.width {
                return f;
            }
        }

        if let Some(sheet_format_properties) = sheet_format_properties {
            return if let Some(f) = sheet_format_properties.default_col_width {
                f
            } else if let Some(f) = sheet_format_properties.base_col_width {
                // defaultColWidth = baseColumnWidth + {margin padding (2 pixels on each side, totalling 4 pixels)} + {gridline (1pixel)}
                (f + 4 + 1) as f64
            } else {
                DEFAULT_CELL_WIDTH
            };
        }

        return DEFAULT_CELL_WIDTH;
    }

    fn cell_height(row_info: RawRow, sheet_format_properties: Option<RawSheetProperties>) -> f64 {
        if let Some(f) = row_info.height {
            return f;
        }

        if let Some(sheet_format_properties) = sheet_format_properties {
            return sheet_format_properties
                .default_row_height
                .unwrap_or(DEFAULT_CELL_HEIGHT);
        }

        return DEFAULT_CELL_HEIGHT;
    }

    fn cell_hidden(
        cell_protection: Option<Protection>,
        col_info: Option<RawCol>,
        row_info: RawRow,
        sheet_format_properties: Option<RawSheetProperties>,
    ) -> bool {
        if let Some(protection) = cell_protection {
            if protection.hidden == Some(true) {
                return true;
            }
        }

        if let Some(b) = row_info.hidden {
            return b;
        }

        if let Some(sheet_format_properties) = sheet_format_properties {
            if let Some(b) = sheet_format_properties.zero_height {
                if b == true {
                    return true;
                }
            }
        }

        if let Some(col_info) = col_info {
            if let Some(b) = col_info.hidden {
                return b;
            }
        }

        return false;
    }
}
