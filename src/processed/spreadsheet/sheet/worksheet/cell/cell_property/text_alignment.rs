use crate::raw::spreadsheet::stylesheet::format::alignment::XlsxAlignment as RawAlignment;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.alignment?view=openxml-3.0.1
///
/// Formatting information pertaining to text alignment in cells.
#[derive(Debug, Clone, PartialEq)]
pub struct TextAlignment {
    /// Specifies the type of horizontal alignment in cells.
    pub horizontal: HorizontalAlignementValue,

    /// Indicates the number of spaces (of the normal style font) of indentation for text in a cell.
    ///
    /// An integer value, where an increment of 1 represents 3 spaces.
    ///
    /// The number of spaces to indent is calculated as following:
    /// ```
    /// Number of spaces to indent = indent value * 3
    /// ```
    pub indent: u64,

    /// A boolean value indicating if the cells justified or distributed alignment should be used on the last line of text.
    // tag: justifyLastLine
    pub justify_last_line: bool,

    /// An integer value indicating whether the reading order (bidirectionality) of the cell is left-to-right, right-to-left, or context dependent.
    /// 0 - Context Dependent - reading order is determined by scanning the text for the first non-whitespace character: if it is a strong right-to-left character, the reading order is right-to-left; otherwise, the reading order left-to-right.
    /// 1 - Left-to-Right- reading order is left-to-right in the cell, as in English.
    /// 2 - Right-to-Left - reading order is right-to-left in the cell, as in Hebrew.
    // tag: readingOrder
    pub reading_order: ReadingOrderValue,

    /// An integer value (used only in a dxf element) to indicate the additional number of spaces of indentation to adjust for text in a cell.
    // tag: relativeIndent
    pub relative_indent: i64,

    /// A boolean value indicating if the displayed text in the cell should be shrunk to fit the cell width.
    ///
    /// Not applicable when a cell contains multiple lines of text.
    pub shrink_to_fit: bool,

    /// Text rotation in cells. Expressed in degrees.
    ///
    /// For 0 - 90, the value represents degrees above horizon.
    /// For 91-180,  the value represents degrees below horizon.
    pub text_rotation: u64,

    /// Vertical alignment in cells.
    pub vertical_alignment: VerticalAlignementValue,

    /// A boolean value indicating if the text in a cell should be line-wrapped within the cell.
    pub wrap_text: bool,
}

impl TextAlignment {
    pub(crate) fn default() -> Self {
        return Self {
            horizontal: HorizontalAlignementValue::General,
            indent: 0,
            justify_last_line: false,
            reading_order: ReadingOrderValue::LeftToRight,
            relative_indent: 0,
            shrink_to_fit: false,
            text_rotation: 0,
            vertical_alignment: VerticalAlignementValue::Bottom,
            wrap_text: false,
        };
    }
    pub(crate) fn from_raw(alignment: Option<RawAlignment>) -> Self {
        let Some(alignment) = alignment else {
            return Self::default();
        };
        return Self {
            horizontal: HorizontalAlignementValue::from_string(alignment.horizontal),
            indent: alignment.indent.unwrap_or(0),
            justify_last_line: alignment.justify_last_line.unwrap_or(false),
            reading_order: ReadingOrderValue::from_index(alignment.reading_order),
            relative_indent: alignment.relative_indent.unwrap_or(0),
            shrink_to_fit: alignment.shrink_to_fit.unwrap_or(false),
            text_rotation: alignment.text_rotation.unwrap_or(0),
            vertical_alignment: VerticalAlignementValue::from_string(alignment.vertical),
            wrap_text: alignment.wrap_text.unwrap_or(false),
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.horizontalalignmentvalues?view=openxml-3.0.1
///
// * Center
// * CenterContinuous
// * Distributed
// * Fill
// * General
// * Justify
// * Left
// * Right
#[derive(Debug, Clone, PartialEq)]
pub enum HorizontalAlignementValue {
    Center,
    CenterContinuous,
    Distributed,
    Fill,
    General,
    Justify,
    Left,
    Right,
}

impl HorizontalAlignementValue {
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::General };
        return match s.as_ref() {
            "center" => Self::Center,
            "centerContinuous" => Self::CenterContinuous,
            "distributed" => Self::Distributed,
            "fill" => Self::Fill,
            "general" => Self::General,
            "justify" => Self::Justify,
            "left" => Self::Left,
            "right" => Self::Right,
            _ => Self::General,
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.verticalalignmentvalues?view=openxml-3.0.1
///
/// * Bottom
/// * Center
/// * Distributed
/// * Justify
/// * Top
#[derive(Debug, Clone, PartialEq)]
pub enum VerticalAlignementValue {
    Bottom,
    Center,
    Distributed,
    Justify,
    Top,
}

impl VerticalAlignementValue {
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::Bottom };
        return match s.as_ref() {
            "bottom" => Self::Bottom,
            "center" => Self::Center,
            "distributed" => Self::Distributed,
            "justify" => Self::Justify,
            "top" => Self::Top,
            _ => Self::Bottom,
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.alignment?view=openxml-3.0.1
///
#[derive(Debug, Clone, PartialEq)]
pub enum ReadingOrderValue {
    ContextDependent,
    LeftToRight,
    RightToLeft,
}

impl ReadingOrderValue {
    pub(crate) fn from_index(s: Option<u64>) -> Self {
        let Some(s) = s else {
            return Self::LeftToRight;
        };
        return match s {
            0 => Self::ContextDependent,
            1 => Self::LeftToRight,
            2 => Self::RightToLeft,
            _ => Self::LeftToRight,
        };
    }
}
