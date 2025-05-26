use std::io::Read;

use quick_xml::events::BytesStart;

use crate::excel::XmlReader;

use super::locks_base::XlsxLocksBase;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.connectionshapelocks?view=openxml-3.0.1
///
///
/// This element specifies all locking properties for a connection shape.
/// These properties inform the generating application about specific properties that have been previously locked and thus should not be changed.
///
/// cxnSpLocks
pub type XlsxConnectionShapeLocks = XlsxLocksBase;

pub(crate) fn load_connection_shape_locks(
    reader: &mut XmlReader<impl Read>,
    e: &BytesStart,
) -> anyhow::Result<XlsxConnectionShapeLocks> {
    return XlsxLocksBase::load(reader, e, b"cxnSpLocks");
}
