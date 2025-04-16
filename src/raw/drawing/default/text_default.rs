use crate::excel::XmlReader;

use super::DefaultBase;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textdefault?view=openxml-3.0.1
///
/// This element defines the default formatting which is applied to text in a document by default.
///
/// Example
/// ```
/// <txDef>
///   <spPr>
///     <solidFill>
///       <schemeClr val="accent2">
///         <shade val="75000"/>
///       </schemeClr>
///     </solidFill>
///   </spPr>
///   <bodyPr rtlCol="0" anchor="ctr"/>
///   <lstStyle>
///     <defPPr algn="ctr">
///       <defRPr/>
///     </defPPr>
///   </lstStyle>
///   <style>
///     <lnRef idx="1">
///       <schemeClr val="accent1"/>
///     </lnRef>
///     <fillRef idx="2">
///       <schemeClr val="accent1"/>
///     </fillRef>
///     <effectRef idx="1">
///       <schemeClr val="accent1"/>
///     </effectRef>
///     <fontRef idx="minor">
///       <schemeClr val="dk1"/>
///     </fontRef>
///   </style>
/// </txDef>
/// ```
pub type TextDefault = DefaultBase;

pub(crate) fn load_text_default(reader: &mut XmlReader) -> anyhow::Result<TextDefault> {
    return DefaultBase::load(reader, b"txDef");
}
