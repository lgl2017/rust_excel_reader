use chrono::FixedOffset;
use chrono::NaiveDateTime;

use crate::helper::a1_address_to_row_col;
use crate::helper::a1_dimension_to_row_col;
use crate::helper::r1c1_address_to_row_col;
use crate::helper::r1c1_dimension_to_row_col;
use crate::helper::string_to_int;

#[cfg(feature = "serde")]
use serde::{Deserialize, Serialize};

/// Hex representation of RGBA (alpha last)
///
/// ex: #88f94eff
pub type HexColor = String;

pub type Text = String;

pub type XlsxSimplePercentage = i64;

/// https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_AdjCoordinate_topic_ID0E14KNB.html
///
/// `ST_AdjCoordinate` defined as a union of the following
/// - `ST_Coordinate` simple type: i64
/// - `ST_GeomGuideName`: String referencing to a geometry guide name
#[derive(Debug, Clone, PartialEq)]
pub enum XlsxAdjustCoordinate {
    GuideName(String),
    Coordinate(i64),
}

impl XlsxAdjustCoordinate {
    pub fn from_string(str: &str) -> Self {
        return if let Some(coordinate) = string_to_int(str) {
            Self::Coordinate(coordinate)
        } else {
            Self::GuideName(str.to_owned())
        };
    }
}

/// https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_AdjAngle_topic_ID0EZWKNB.html
/// `ST_AdjAngle` defined as a union of the following
/// - `ST_Angle` simple type: i64
/// - `ST_GeomGuideName`: String referencing to a geometry guide name
#[derive(Debug, Clone, PartialEq)]
pub enum XlsxAdjustAngle {
    GuideName(String),
    Angle(i64),
}

impl XlsxAdjustAngle {
    pub fn from_string(str: &str) -> Self {
        return if let Some(angle) = string_to_int(str) {
            Self::Angle(angle)
        } else {
            Self::GuideName(str.to_owned())
        };
    }
}

/// row, col: 1 based index
#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Default, PartialEq, Eq, Hash, Ord, PartialOrd, Copy, Clone)]
pub struct Coordinate {
    pub row: u64,
    pub col: u64,
}

impl Coordinate {
    pub fn from_point(point: (u64, u64)) -> Self {
        Self {
            row: point.0,
            col: point.1,
        }
    }

    pub fn from_a1(a1_address: &[u8]) -> Option<Self> {
        if let Ok((Some(row), Some(col))) = a1_address_to_row_col(a1_address) {
            return Some(Self { row, col });
        }
        return None;
    }

    pub fn from_r1c1(r1c1: &str) -> Option<Self> {
        if let Ok(Some(coordinate)) = r1c1_address_to_row_col(r1c1) {
            return Some(Self {
                row: coordinate.0,
                col: coordinate.1,
            });
        }
        return None;
    }
}

#[cfg_attr(feature = "serde", derive(Serialize, Deserialize))]
#[derive(Debug, Default, PartialEq, Eq, Hash, Ord, PartialOrd, Copy, Clone)]
pub struct Dimension {
    pub start: Coordinate,
    pub end: Coordinate,
}

impl Dimension {
    pub(crate) fn from_a1(a1_address: &[u8]) -> Option<Self> {
        if let Ok((start, end)) = a1_dimension_to_row_col(a1_address) {
            return Some(Self {
                start: Coordinate::from_point(start),
                end: Coordinate::from_point(end),
            });
        }
        return None;
    }

    pub(crate) fn from_r1c1(r1c1: &str) -> Option<Self> {
        if let Ok((start, end)) = r1c1_dimension_to_row_col(r1c1) {
            return Some(Self {
                start: Coordinate::from_point(start),
                end: Coordinate::from_point(end),
            });
        }
        return None;
    }
}

#[derive(Debug, PartialEq, Eq, Hash, Copy, Clone)]
pub struct XlsxDatetime {
    pub datetime: NaiveDateTime,
    pub offset: Option<FixedOffset>,
}
