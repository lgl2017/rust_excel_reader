/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.alphafloor?view=openxml-3.0.1
/// when present, Alpha (opacity) values less than 100% are changed to zero.
/// In other words, anything partially transparent becomes fully transparent.
pub type XlsxAlphaFloor = bool;
