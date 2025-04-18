pub mod border;
pub mod fill;
pub mod font;
pub mod hyperlink;
pub mod numbering_format;
pub mod text_alignment;

use border::Border;
use fill::Fill;
use font::Font;
use hyperlink::Hyperlink;
use numbering_format::NumberingFormat;
use text_alignment::TextAlignment;

use crate::raw::{
    drawing::scheme::color_scheme::XlsxColorScheme,
    spreadsheet::{
        sheet::{
            sheet_format_properties::XlsxSheetFormatProperties,
            worksheet::{cell::XlsxCell, column_information::XlsxColumnInformation, row::XlsxRow},
        },
        stylesheet::{
            format::{alignment::XlsxAlignment, protection::XlsxCellProtection},
            XlsxStyleSheet,
        },
    },
};

static DEFAULT_CELL_WIDTH: f64 = 8.43;
static DEFAULT_BEST_FIT: bool = false;
static DEFAULT_CELL_HEIGHT: f64 = 15.0;
static DEFAULT_DY_DESCENT: f64 = 0.2;
static DEFAULT_HIDDEN: bool = false;
static DEFAULT_SHOW_PHONETIC: bool = true;

#[derive(Debug, Clone, PartialEq)]
pub struct CellProperty {
    pub width: f64,

    /// Flag indicating if the column width should automatically resize
    pub width_best_fit: bool,

    pub height: f64,

    /// vertical distance in pixels from the bottom of a cell in a row to the typographical baseline of its content
    pub dy_descent: f64,

    pub hidden: bool,

    pub show_phonetic: bool,

    pub hyperlink: Option<Hyperlink>,

    // styles
    pub alignment: TextAlignment,
    pub font: Font,
    pub border: Border,
    pub fill: Fill,
    pub numbering_format: NumberingFormat,
}

impl CellProperty {
    pub(crate) fn default() -> Self {
        return Self {
            width: DEFAULT_CELL_WIDTH,
            width_best_fit: DEFAULT_BEST_FIT,
            height: DEFAULT_CELL_HEIGHT,
            dy_descent: DEFAULT_DY_DESCENT,
            hidden: DEFAULT_HIDDEN,
            show_phonetic: DEFAULT_SHOW_PHONETIC,
            hyperlink: None,
            alignment: TextAlignment::default(),
            font: Font::default(),
            border: Border::default(),
            fill: Fill::default(),
            numbering_format: NumberingFormat::default(),
        };
    }

    pub(crate) fn from_raw(
        cell: XlsxCell,
        row_info: XlsxRow,
        col_info: Option<XlsxColumnInformation>,
        fill_id: Option<u64>,
        font_id: Option<u64>,
        border_id: Option<u64>,
        numbering_format_id: Option<u64>,
        alignment: Option<XlsxAlignment>,
        protection: Option<XlsxCellProtection>,
        hyperlink: Option<Hyperlink>,
        sheet_format_properties: Option<XlsxSheetFormatProperties>,
        stylesheet: XlsxStyleSheet,
        color_scheme: Option<XlsxColorScheme>,
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
            dy_descent: Self::get_dy_descent(row_info, sheet_format_properties),
            hidden,
            show_phonetic,
            hyperlink,
            alignment: TextAlignment::from_raw(alignment),
            font,
            border,
            fill,
            numbering_format,
        };
    }

    fn get_width_best_fit(col_info: Option<XlsxColumnInformation>) -> bool {
        let Some(col_info) = col_info else {
            return DEFAULT_BEST_FIT;
        };

        return col_info.best_fit.unwrap_or(DEFAULT_BEST_FIT);
    }

    fn get_dy_descent(
        row: XlsxRow,
        sheet_format_properties: Option<XlsxSheetFormatProperties>,
    ) -> f64 {
        if let Some(d) = row.dy_descent {
            return d;
        };

        if let Some(sheet_format_properties) = sheet_format_properties {
            if let Some(d) = sheet_format_properties.dy_descent {
                return d;
            }
        };

        return DEFAULT_DY_DESCENT;
    }

    fn get_numbering_format(
        numbering_format: Option<u64>,
        stylesheet: XlsxStyleSheet,
    ) -> NumberingFormat {
        if let Some(id) = numbering_format {
            let raw = stylesheet.get_num_format(id);
            return NumberingFormat::from_raw(raw);
        }

        return NumberingFormat::default();
    }

    fn get_font(
        font_id: Option<u64>,
        stylesheet: XlsxStyleSheet,
        color_scheme: Option<XlsxColorScheme>,
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
        stylesheet: XlsxStyleSheet,
        color_scheme: Option<XlsxColorScheme>,
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
        stylesheet: XlsxStyleSheet,
        color_scheme: Option<XlsxColorScheme>,
    ) -> Fill {
        if let Some(id) = fill_id {
            if let Ok(id) = TryInto::<usize>::try_into(id) {
                let raw = stylesheet.get_fill(id);
                return Fill::from_raw(raw, stylesheet.colors.clone(), color_scheme);
            }
        }

        return Fill::default();
    }

    fn show_phonetic(
        cell: XlsxCell,
        col_info: Option<XlsxColumnInformation>,
        row_info: XlsxRow,
    ) -> bool {
        return if let Some(b) = cell.show_phonetic {
            b
        } else if let Some(b) = row_info.show_phonetic {
            b
        } else if let Some(col_info) = col_info {
            if let Some(b) = col_info.show_phonetic {
                b
            } else {
                DEFAULT_SHOW_PHONETIC
            }
        } else {
            DEFAULT_SHOW_PHONETIC
        };
    }

    fn cell_width(
        col_info: Option<XlsxColumnInformation>,
        sheet_format_properties: Option<XlsxSheetFormatProperties>,
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

    fn cell_height(
        row_info: XlsxRow,
        sheet_format_properties: Option<XlsxSheetFormatProperties>,
    ) -> f64 {
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
        cell_protection: Option<XlsxCellProtection>,
        col_info: Option<XlsxColumnInformation>,
        row_info: XlsxRow,
        sheet_format_properties: Option<XlsxSheetFormatProperties>,
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

        return DEFAULT_HIDDEN;
    }
}
