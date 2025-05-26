#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::{st_types::emu_to_pt, text::paragraph::tab_stop::XlsxTabStop};

use super::tab_alignment_values::TextTabAlignmentValues;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.tabstop?view=openxml-3.0.1
///
/// This element specifies a single tab stop to be used on a line of text when there are one or more tab characters present within the text.
///
/// Example:
/// ```
/// <a:tabLst>
///     <a:tab pos="2292350" algn="l"/>
///     <a:tab pos="2627313" algn="l"/>
///     <a:tab pos="2743200" algn="l"/>
///     <a:tab pos="2974975" algn="l"/>
/// </a:tabLst>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]

pub struct TabStop {
    /// Specifies the alignment that is to be applied to text using this tab stop.
    /// If this attribute is omitted then the application default for the generating application.
    ///
    /// Possible Values: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.texttabalignmentvalues?view=openxml-3.0.1
    pub alignment: TextTabAlignmentValues,

    /// Specifies the position of the tab stop relative to the left margin.
    /// If this attribute is omitted then the application default for tab stops is used.
    pub position: f64,
}

impl TabStop {
    pub(crate) fn from_raw(raw: XlsxTabStop) -> Self {
        return Self {
            alignment: TextTabAlignmentValues::from_string(raw.alignment),
            position: emu_to_pt(raw.position.unwrap_or(914400)), // 1 inch
        };
    }
}
