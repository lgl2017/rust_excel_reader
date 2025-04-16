use super::{phonetic_properties::PhoneticProperties, phonetic_run::PhoneticRun};
use crate::{
    common_types::Text, processed::spreadsheet::sheet::worksheet::cell_property::font::Font,
};

/// Example:
/// ```
/// // shared string
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
#[derive(Debug, Clone, PartialEq)]
pub struct RichText {
    pub phonetic_properties: Option<PhoneticProperties>,
    pub phonetic_runs: Option<Vec<PhoneticRun>>,
    pub runs: Vec<RichTextRun>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct RichTextRun {
    pub font: Font,
    pub text: Text,
}
