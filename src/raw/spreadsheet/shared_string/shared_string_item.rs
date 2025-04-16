use crate::{excel::XmlReader, raw::spreadsheet::string_item::StringItem};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.sharedstringitem?view=openxml-3.0.1
///
/// This element is the representation of an individual string in the Shared String table.
///
/// If the string is just a simple string with formatting applied at the cell level, then the String Item (si) should contain a single text element used to express the string.
/// However, if the string in the cell is more complex - i.e., has formatting applied at the character level - then the string item shall consist of multiple rich text runs which collectively are used to express the string.
///
///
/// Example:
/// ```
/// <si>
///     <t>課きく　毛こ</t>
///     <rPh sb="0" eb="1">
///         <t>カ</t>
///     </rPh>
///     <rPh sb="4" eb="5">
///        <t>ケ</t>
///     </rPh>
///     <phoneticPr fontId="1"/>
/// </si>
/// <si>
///     <r>
///         <rPr>
///             <sz val="10" />
///             <color indexed="8" />
///             <rFont val="Helvetica Neue" />
///         </rPr>
///         <t>123</t>
///     </r>
///     <r>
///         <rPr>
///             <b val="1" />
///             <sz val="10" />
///             <color indexed="8" />
///             <rFont val="Helvetica Neue" />
///         </rPr>
///         <t>4</t>
///     </r>
/// </si>
/// ```
pub type SharedStringItem = StringItem;

pub(crate) fn load_shared_string_item(reader: &mut XmlReader) -> anyhow::Result<SharedStringItem> {
    return StringItem::load(reader, b"si");
}
