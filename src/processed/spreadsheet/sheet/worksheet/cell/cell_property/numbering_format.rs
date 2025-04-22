#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::spreadsheet::stylesheet::{
    format::numbering_format::get_builtin_format_code, XlsxStyleSheet,
};

static DEFAULT_FORMAT_CODE: &str = "general";
static DEFAULT_FORMAT_ID: u64 = 0;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-3.0.1
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
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

    pub(crate) fn from_id(num_format_id: Option<u64>, stylesheet: XlsxStyleSheet) -> Self {
        let Some(num_format_id) = num_format_id else {
            return Self::default();
        };
        let format = stylesheet.get_num_format(num_format_id);
        let mut format_code = get_builtin_format_code(num_format_id);

        if let Some(format) = format {
            if let Some(code) = format.format_code {
                format_code = Some(code)
            }
        };

        return Self {
            format_code,
            format_id: num_format_id,
        };
    }
}
