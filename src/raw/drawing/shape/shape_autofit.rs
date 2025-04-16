/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapeautofit?view=openxml-3.0.1
/// This element specifies that a shape should be auto-fit to fully contain the text described within it.
/// Auto-fitting is when text within a shape is scaled in order to contain all the text inside.
/// If this element is omitted, then noAutofit or auto-fit off is implied.
///
/// Example:
/// ```
/// <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="50800" tIns="50800" rIns="50800" bIns="50800" numCol="1" spcCol="38100" rtlCol="0" anchor="ctr" upright="0">
///     <a:spAutoFit />
/// </a:bodyPr>
/// ```
pub type ShapeAutofit = bool;
