use crate::{excel::XmlReader, raw::spreadsheet::string_item::StringItem};

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.inlinestring?view=openxml-3.0.1
///
/// This element allows for strings to be expressed directly in the cell definition instead of implementing the shared string table.
///
/// Example
/// ```
/// <c r="A1" t="inlineStr">
///     <is><t>This is inline string example</t></is>
/// </c>
/// ```
///
/// is (Rich Text Inline)
pub type InlineString = StringItem;

pub(crate) fn load_inline_string(reader: &mut XmlReader) -> anyhow::Result<InlineString> {
    return StringItem::load(reader, b"is");
}
