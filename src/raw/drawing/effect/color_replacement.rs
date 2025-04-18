use crate::raw::drawing::color::XlsxColorEnum;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.colorreplacement?view=openxml-3.0.1
/// specifies a solid color replacement value.
/// All effect colors are changed to a fixed color. Alpha values are unaffected.
// tag: clrRepl
pub type XlsxColorReplacement = XlsxColorEnum;
