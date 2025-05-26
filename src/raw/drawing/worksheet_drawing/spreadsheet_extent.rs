use crate::raw::drawing::shape::extents::XlsxExtents;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.extent?view=openxml-3.0.1
///
/// This element describes the length and width properties for how far a drawing element should extend for.
///
/// Example:
/// ```
/// <xdr:ext cx="2426208" cy="978408"/>
/// ```
pub type XlsxSpreadsheetExtent = XlsxExtents;
