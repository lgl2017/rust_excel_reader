#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

use crate::common_types::Text;

use super::{phonetic_properties::PhoneticProperties, phonetic_run::PhoneticRun};

/// Example:
///
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
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub struct PlainText {
    pub phonetic_properties: Option<PhoneticProperties>,
    pub phonetic_runs: Option<Vec<PhoneticRun>>,
    pub text: Text,
}
