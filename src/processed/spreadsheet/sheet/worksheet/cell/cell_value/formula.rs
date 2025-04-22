#[cfg(feature = "serde")]
use serde::Serialize;

#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct Formula {
    pub formula: String,
    pub last_calculated_value: Option<String>,
}
