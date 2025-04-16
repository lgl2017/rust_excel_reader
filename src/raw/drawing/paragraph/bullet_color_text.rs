/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.bulletcolortext?view=openxml-3.0.1
/// This element specifies that the color of the bullets for a paragraph should be of the same color as the text run within which each bullet is contained.
///
/// Example:
/// ```
/// <a:pPr …>
///     <a:buClrTx />
/// </a:pPr>
/// ```
// tag: buClrTx
pub type BulletColorText = bool;
