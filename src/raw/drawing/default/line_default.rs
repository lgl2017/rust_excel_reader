use crate::excel::XmlReader;

use super::DefaultBase;

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
pub type LineDefault = DefaultBase;

pub(crate) fn load_line_default(reader: &mut XmlReader) -> anyhow::Result<LineDefault> {
    return DefaultBase::load(reader, b"lnDef");
}
