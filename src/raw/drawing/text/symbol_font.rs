use super::font::text_font_type::XlsxTextFontType;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.symbolfont?view=openxml-3.0.1
///
/// This element specifies that a symbol font be used for a specific run of text.
///
/// Example
/// ```
/// <a:rPr …>
///     <a:sym typeface="Sample Font"/>
/// </a:rPr>
/// ```
// tag: sym
pub type XlsxSymbolFont = XlsxTextFontType;
