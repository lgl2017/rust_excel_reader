use std::io::Read;

use quick_xml::events::BytesStart;

use crate::excel::XmlReader;

use super::locks_base::XlsxLocksBase;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.picturelocks?view=openxml-3.0.1
///
/// picLocks (Picture Locks)
pub type XlsxPictureLocks = XlsxLocksBase;

pub(crate) fn load_picture_locks(
    reader: &mut XmlReader<impl Read>,
    e: &BytesStart,
) -> anyhow::Result<XlsxPictureLocks> {
    return XlsxLocksBase::load(reader, e, b"picLocks");
}
