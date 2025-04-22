use std::fmt;

use anyhow::bail;

#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// https://msdn.microsoft.com/en-us/library/office/ff839168.aspx
///
/// Errors that can appear as a value in a worksheet cell
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
pub enum CellErrorType {
    /// Division by 0 error
    Div0,
    /// Unavailable value error
    NA,
    /// Invalid name error
    Name,
    /// Null value error
    Null,
    /// Number error
    Num,
    /// Invalid cell reference error
    Ref,
    /// Value error
    Value,
    /// Getting data
    GettingData,
}

impl fmt::Display for CellErrorType {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> Result<(), fmt::Error> {
        match *self {
            CellErrorType::Div0 => write!(f, "#DIV/0!"),
            CellErrorType::NA => write!(f, "#N/A"),
            CellErrorType::Name => write!(f, "#NAME?"),
            CellErrorType::Null => write!(f, "#NULL!"),
            CellErrorType::Num => write!(f, "#NUM!"),
            CellErrorType::Ref => write!(f, "#REF!"),
            CellErrorType::Value => write!(f, "#VALUE!"),
            CellErrorType::GettingData => write!(f, "#DATA!"),
        }
    }
}

impl CellErrorType {
    pub(crate) fn from_string(s: &str) -> anyhow::Result<Self> {
        match s {
            "#DIV/0!" => Ok(CellErrorType::Div0),
            "#N/A" => Ok(CellErrorType::NA),
            "#NAME?" => Ok(CellErrorType::Name),
            "#NULL!" => Ok(CellErrorType::Null),
            "#NUM!" => Ok(CellErrorType::Num),
            "#REF!" => Ok(CellErrorType::Ref),
            "#VALUE!" => Ok(CellErrorType::Value),
            _ => bail!("Unkown cell error value."),
        }
    }
}
