use crate::helper::string_to_int;

/// Represent the following types:
/// * [ST_TextFontSize](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_TextFontSize_topic_ID0EPPQOB.html)
///     - a minimum value of greater than or equal to 100 (1pt).
///     - a maximum value of less than or equal to 400000 (4000pt).
///
/// * [ST_TextNonNegativePoint](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_TextNonNegativePo_topic_ID0ERISOB.html)
///     - a minimum value of greater than or equal to 0 (0pt).
///     - a maximum value of less than or equal to 400000 (4000pt).
///
/// * [ST_TextSpacingPoint](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_TextSpacingPoint_topic_ID0EYCUOB.html)
///     - a minimum value of greater than or equal to 0.
///     - a maximum value of less than or equal to 158400.
///
/// This type specifies the size of any text in hundredths of a point.
pub type STPositiveTextPoint = u64;

/// Represent the following types:
/// * [ST_TextPoint](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_TextPoint_topic_ID0E6NSOB.html)
///     - a minimum value of greater than or equal to -400000 (-4000pt).
///     - a maximum value of less than or equal to 400000 (4000pt).
///
/// This type specifies the size of any text in hundredths of a point.
pub type STTextPoint = u64;

/// Converting the following type to points.
///
/// * [ST_TextFontSize](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_TextFontSize_topic_ID0EPPQOB.html)
///     - a minimum value of greater than or equal to 100 (1pt).
///     - a maximum value of less than or equal to 400000 (4000pt).
///
/// * [ST_TextNonNegativePoint](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_TextNonNegativePo_topic_ID0ERISOB.html)
///     - a minimum value of greater than or equal to 0 (0pt).
///     - a maximum value of less than or equal to 400000 (4000pt).
///
/// * [ST_TextSpacingPoint](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_TextSpacingPoint_topic_ID0EYCUOB.html)
///     - a minimum value of greater than or equal to 0.
///     - a maximum value of less than or equal to 158400.
/// * [ST_TextPoint](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_TextPoint_topic_ID0E6NSOB.html)
///     - a minimum value of greater than or equal to -400000 (-4000pt).
///     - a maximum value of less than or equal to 400000 (4000pt).
///
/// Ex: 100 -> 1.0 pt
#[allow(dead_code)]
pub(crate) fn st_text_point_to_pt(int: i64) -> f64 {
    return (int as f64) / 100.0;
}

/// Represent the following types:
/// * [ST_FixedPercentage](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_FixedPercentage_topic_ID0E6NSNB.html)
/// * [ST_Percentage](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_Percentage_topic_ID0EY3XNB.html#topic_ID0EY3XNB)
///
/// This simple type represents a fixed percentage in 1000ths of a percent.
///
/// Example:
///
/// `100%`: `100000` in `XlsxFixedPercentage`
pub type STPercentage = i64;

/// Represent the following types:
/// * [ST_PositiveFixedPercentage](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_PositiveFixedPerc_topic_ID0EIV1NB.html)
///
/// This simple type represents a positive fixed percentage in 1000ths of a percent.
pub type STPositivePercentage = u64;

/// Converting [ST_Percentage](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_Percentage_topic_ID0EY3XNB.html) Integer to float.
///
/// Ex: 56,000 => 0.56
pub(crate) fn st_percentage_to_float(int: i64) -> f64 {
    return (int as f64) / 1000.0 / 100.0;
}

/// Represent the following types:
/// * [ST_FixedAngle](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_FixedAngle_topic_ID0EFGSNB.html)
/// * [ST_Angle](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_Angle_topic_ID0EXLLNB.html#topic_ID0EXLLNB)
///
/// This simple type represents an angle in 60,000ths of a degree.
///
/// Example:
///
/// `90 degree`: `5400000` in `STAngle`
pub type STAngle = i64;

/// Represent the following types:
/// * [ST_PositiveFixedAngle](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_PositiveFixedAngl_topic_ID0ECJ1NB.html)
///
/// This simple type represents a positive angle in 60000ths of a degree. Range from [0, 360 degrees).
pub type STPositiveAngle = u64;

/// Converting [ST_FixedAngle](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_FixedAngle_topic_ID0EFGSNB.html) to float.
///
/// Ex: 5400000 => 90
pub(crate) fn st_angle_to_degree(int: i64) -> f64 {
    return (int as f64) / 60000.0;
}

/// Represent the following types:
/// * [ST_PositiveCoordinate](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_PositiveCoordinat_topic_ID0EUHZNB.html)
/// * [ST_PositiveCoordinate32](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_PositiveCoordinat_topic_ID0EWB1NB.html)
///
/// This simple type represents a positive position or length in EMUs.
pub type STPositiveCoordinate = u64;

/// Represent the following types:
/// * [ST_Coordinate](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_Coordinate_topic_ID0E4DQNB.html)
/// * [ST_Coordinate32](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_Coordinate32_topic_ID0EXBRNB.html)
///
/// This simple type represents a one dimensional position or length in EMUs.
pub type STCoordinate = i64;

/// Convert EMU to points
///
/// English Metric Unit (EMU): 1/360,000 of a centimeter, with 914,400 EMUs per inch and 12,700 EMUs per point.
///
/// - 1 EMU = 1/360,000 cm
/// - 1 inch = 914,400 EMUs
/// - 1 point = 12,700 EMUs
#[allow(dead_code)]
pub(crate) fn emu_to_pt(emu: i64) -> f64 {
    return (emu as f64) / 12700.0;
}

/// https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_AdjCoordinate_topic_ID0E14KNB.html
///
/// `ST_AdjCoordinate` defined as a union of the following
/// - `ST_Coordinate` simple type: i64. Represents a one dimensional position or length in EMUs: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/4f890b34-61b8-4d22-beb7-77ac953e66a8)
/// - `ST_GeomGuideName`: String referencing to a geometry guide name: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_GeomGuideName_topic_ID0ELGTNB.html
#[derive(Debug, Clone, PartialEq)]
pub enum STAdjustCoordinate {
    /// Name refer to a geometry guide name (gd)
    ///
    /// Example:
    /// ```
    /// <a:custGeom>
    ///   <a:avLst/>
    ///   <a:gdLst>
    ///     <a:gd name="myGuide" fmla="*/ h 2 3"/>
    ///   </a:gdLst>
    ///   <a:ahLst/>
    ///   <a:cxnLst/>
    ///   <a:rect l="0" t="0" r="0" b="0"/>
    ///   <a:pathLst>
    ///     <a:path w="1705233" h="679622">
    ///       <a:moveTo>
    ///         <a:pt x="0" y="myGuide"/>
    ///       </a:moveTo>
    ///       <a:lnTo>
    ///         <a:pt x="1705233" y="myGuide"/>
    ///       </a:lnTo>
    ///       <a:lnTo>
    ///         <a:pt x="852616" y="0"/>
    ///       </a:lnTo>
    ///       <a:close/>
    ///     </a:path>
    ///   </a:pathLst>
    /// </a:custGeom>
    /// ```
    GuideName(String),
    Coordinate(STCoordinate),
}

impl STAdjustCoordinate {
    pub fn from_string(str: &str) -> Self {
        return if let Some(coordinate) = string_to_int(str) {
            Self::Coordinate(coordinate)
        } else {
            Self::GuideName(str.to_owned())
        };
    }
}

/// https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_AdjAngle_topic_ID0EZWKNB.html
///
/// `ST_AdjAngle` defined as a union of the following
/// - `ST_Angle` simple type: i64: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_Angle_topic_ID0EXLLNB.html#topic_ID0EXLLNB
/// - `ST_GeomGuideName`: String referencing to a geometry guide name: https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_GeomGuideName_topic_ID0ELGTNB.html
#[derive(Debug, Clone, PartialEq)]
pub enum STAdjustAngle {
    GuideName(String),
    Angle(STAngle),
}

impl STAdjustAngle {
    pub fn from_string(str: &str) -> Self {
        return if let Some(angle) = string_to_int(str) {
            Self::Angle(angle)
        } else {
            Self::GuideName(str.to_owned())
        };
    }
}
