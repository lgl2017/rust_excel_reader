/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sheetprotection?view=openxml-3.0.1
///
/// This collection expresses the sheet protection options to enforce when the sheet is protected.
///
/// Example:
/// ```
/// <sheetProtection sheet="1" objects="1" scenarios="1" formatCells="0"  selectLockedCells="1"/>
/// ```

#[derive(Debug, Clone, PartialEq)]
pub struct XlsxSheetProtection {
    //     AlgorithmName
    // Cryptographic Algorithm Name
    // Represents the following attribute in the schema: algorithmName
    // AutoFilter
    // AutoFilter Locked
    // Represents the following attribute in the schema: autoFilter
}
