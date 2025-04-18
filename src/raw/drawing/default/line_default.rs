use crate::excel::XmlReader;

use super::XlsxDefaultBase;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linedefault?view=openxml-3.0.1
///
/// This element defines a default line that is used within a document.
///
/// Example
/// ```
/// <lnDef>
///   <spPr/>
///   <bodyPr/>
///   <lstStyle/>
///   <style>
///     <lnRef idx="1">
///       <schemeClr val="accent2"/>
///     </lnRef>
///     <fillRef idx="0">
///       <schemeClr val="accent2"/>
///     </fillRef>
///     <effectRef idx="0">
///       <schemeClr val="accent2"/>
///     </effectRef>
///     <fontRef idx="minor">
///       <schemeClr val="tx1"/>
///     </fontRef>
///   </style>
/// </lnDef>
/// ```
pub type XlsxLineDefault = XlsxDefaultBase;

pub(crate) fn load_line_default(reader: &mut XmlReader) -> anyhow::Result<XlsxLineDefault> {
    return XlsxDefaultBase::load(reader, b"lnDef");
}
