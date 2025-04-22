#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

use super::{phonetic_properties::PhoneticProperties, phonetic_run::PhoneticRun};
use crate::{
    common_types::Text, processed::spreadsheet::sheet::worksheet::cell::cell_property::font::Font,
};

/// Example:
/// ```
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
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct RichText {
    pub phonetic_properties: Option<PhoneticProperties>,
    pub phonetic_runs: Option<Vec<PhoneticRun>>,
    pub runs: Vec<RichTextRun>,
}

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct RichTextRun {
    pub font: Font,
    pub text: Text,
}
