/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetview?view=openxml-3.0.1
///
/// A single sheet view definition.
/// When more than one sheet view is defined in the file, it means that when opening the workbook, each sheet view corresponds to a separate window within the spreadsheet application, where each window is showing the particular sheet containing the same workbookViewId value, the last sheetView definition is loaded, and the others are discarded.
/// Example
/// ```
/// <sheetViews>
///   <sheetView tabSelected="1" workbookViewId="0">
///     <pane xSplit="2310" ySplit="2070" topLeftCell="C1" activePane="bottomRight"/>
///     <selection/>
///     <selection pane="bottomLeft" activeCell="A6" sqref="A6"/>
///     <selection pane="topRight" activeCell="C1" sqref="C1"/>
///     <selection pane="bottomRight" activeCell="E13" sqref="E13"/>
///   </sheetView>
/// </sheetViews>
/// ```
/// sheetView (Worksheet View)
#[derive(Debug, Clone, PartialEq)]
pub struct SheetView {
    // extLst (Future Feature Data Storage Area) Not supported

    // Child Elements
    // pane (View Pane)	§18.3.1.66
    // pivotSelection (PivotTable Selection)	§18.3.1.69
    // selection (Selection)
}
