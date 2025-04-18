/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bulletsizetext?view=openxml-3.0.1
///
/// This element specifies that the size of the bullets for a paragraph should be of the same point size as the text run within which each bullet is contained.
///
/// Example:
/// ```
/// <a:pPr …>
///     <a:buSzTx />
/// </a:pPr>
/// ```
// tag: buSzTx
pub type XlsxBulletSizeText = bool;
