#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

use crate::raw::spreadsheet::stylesheet::format::numbering_format::{
    get_builtin_format_code, XlsxNumberingFormat,
};

static DEFAULT_FORMAT_CODE: &str = "general";
static DEFAULT_FORMAT_ID: u64 = 0;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct NumberingFormat {
    pub format_code: Option<String>,
    pub format_id: u64,
}

impl NumberingFormat {
    pub(crate) fn default() -> Self {
        return Self {
            format_code: Some(DEFAULT_FORMAT_CODE.to_string()),
            format_id: DEFAULT_FORMAT_ID,
        };
    }

    pub(crate) fn from_raw(format: Option<XlsxNumberingFormat>) -> Self {
        let Some(format) = format else {
            return Self::default();
        };
        let Some(num_format_id) = format.num_fmt_id else {
            return Self::default();
        };

        let format_code = if let Some(code) = format.format_code {
            Some(code)
        } else {
            get_builtin_format_code(num_format_id)
        };

        return Self {
            format_code,
            format_id: num_format_id,
        };
    }
}
