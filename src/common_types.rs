use chrono::DateTime;
use chrono::FixedOffset;
use chrono::NaiveDateTime;

use crate::helper::a1_address_to_row_col;
use crate::helper::a1_dimension_to_row_col;
use crate::helper::r1c1_address_to_row_col;
use crate::helper::r1c1_dimension_to_row_col;

#[cfg(feature = "serde")]
use serde::Serialize;

/// Hex representation of RGBA (alpha last)
///
/// ex: #88f94eff
pub type HexColor = String;

pub type Text = String;

/// row, col: 1 based index
#[cfg_attr(feature = "serde", derive(Serialize))]
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

#[cfg_attr(feature = "serde", derive(Serialize))]
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

impl XlsxDatetime {
    /// Converting Attributes string to datetime
    pub(crate) fn from_string(str: &str) -> Option<Self> {
        // with time zone: YYYY-MM-DDThh:mm:ssZ
        if let Ok(date_time) = DateTime::parse_from_rfc3339(str) {
            return Some(Self {
                datetime: date_time.naive_utc(),
                offset: Some(date_time.offset().to_owned()),
            });
        }

        // without time zone: YYYY-MM-DDThh:mm:ss
        if let Ok(naive_date_time) = NaiveDateTime::parse_from_str(&str, "%Y-%m-%dT%H:%M:%S") {
            return Some(Self {
                datetime: naive_date_time,
                offset: None,
            });
        };

        return None;
    }
}
