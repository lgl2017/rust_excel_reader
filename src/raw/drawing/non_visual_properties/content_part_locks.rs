use std::io::Read;

use quick_xml::events::BytesStart;

use crate::excel::XmlReader;

use super::locks_base::XlsxLocksBase;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.drawing.contentpartlocks?view=openxml-3.0.1
///
/// cpLocks
pub type XlsxContentPartLocks = XlsxLocksBase;

pub(crate) fn load_content_part_locks(
    reader: &mut XmlReader<impl Read>,
    e: &BytesStart,
) -> anyhow::Result<XlsxContentPartLocks> {
    return XlsxLocksBase::load(reader, e, b"cpLocks");
}
