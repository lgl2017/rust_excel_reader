use crate::raw::drawing::font::text_font_type::XlsxTextFontType;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bulletfont?view=openxml-3.0.1
/// This element specifies the font to be used on bullet characters within a given paragraph.
///
/// Example
/// ```
/// <a:pPr …>
///     <a:buFont typeface="Calibri"/>
///     <a:buChar char="g"/>
/// </a:pPr>
/// ```
// tag: buFont
pub type XlsxBulletFont = XlsxTextFontType;
