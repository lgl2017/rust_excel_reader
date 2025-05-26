#[cfg(feature = "serde")]
use serde::Serialize;

use crate::raw::drawing::text::paragraph::auto_numbered_bullet::XlsxAutoNumberedBullet;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.autonumberedbullet?view=openxml-3.0.1
///
/// This element specifies that automatic numbered bullet points should be applied to a paragraph.
/// These are not just numbers used as bullet points but instead automatically assigned numbers that are based on both buAutoNum attributes and paragraph level.
///
/// Example:
/// ```
/// <a:pPr …>
///     <a:buAutoNum type="arabicPeriod"/>
/// </a:pPr>
/// ```
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
pub struct AutoNumberedBullet {
    /// Specifies the number that starts a given sequence of automatically numbered bullets.
    ///
    /// When the numbering is alphabetical, the number should map to the appropriate letter.
    /// For instance 1 maps to 'a', 2 to 'b' and so on.
    /// If the numbers are larger than 26, then multiple letters should be used. For instance 27 should be represented as 'aa' and similarly 53 should be 'aaa'.
    ///
    /// value range:
    /// - minimum value of greater than or equal to 1.
    /// - maximum value of less than or equal to 32767.
    pub start_number: u64,

    /// Specifies the numbering scheme that is to be used.
    ///
    /// This allows for the describing of formats other than strictly numbers.
    /// For instance, a set of bullets can be represented by a series of Roman numerals instead of the standard 1,2,3,etc. number set.
    pub bullet_type: TextAutoNumberSchemeValues,
}

impl AutoNumberedBullet {
    pub(crate) fn from_raw(raw: XlsxAutoNumberedBullet) -> Self {
        let start_number = if let Some(s) = raw.clone().start_at {
            if (1..=32767).contains(&s) {
                s
            } else {
                1
            }
        } else {
            1
        };

        return Self {
            start_number,
            bullet_type: TextAutoNumberSchemeValues::from_string(raw.clone().r#type),
        };
    }
}

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textautonumberschemevalues?view=openxml-3.0.1
///
/// * AlphaLowerCharacterParenBoth
/// * AlphaLowerCharacterParenR
/// * AlphaLowerCharacterPeriod
/// * AlphaUpperCharacterParenBoth
/// * AlphaUpperCharacterParenR
/// * AlphaUpperCharacterPeriod
/// * Arabic1Minus
/// * Arabic2Minus
/// * ArabicDoubleBytePeriod
/// * ArabicDoubleBytePlain
/// * ArabicParenBoth
/// * ArabicParenR
/// * ArabicPeriod
/// * ArabicPlain
/// * CircleNumberDoubleBytePlain
/// * CircleNumberWingdingsBlackPlain
/// * CircleNumberWingdingsWhitePlain
/// * EastAsianJapaneseDoubleBytePeriod
/// * EastAsianJapaneseKoreanPeriod
/// * EastAsianJapaneseKoreanPlain
/// * EastAsianSimplifiedChinesePeriod
/// * EastAsianSimplifiedChinesePlain
/// * EastAsianTraditionalChinesePeriod
/// * EastAsianTraditionalChinesePlain
/// * Hebrew2Minus
/// * HindiAlpha1Period
/// * HindiAlphaPeriod
/// * HindiNumberParenthesisRight
/// * HindiNumPeriod
/// * RomanLowerCharacterParenBoth
/// * RomanLowerCharacterParenR
/// * RomanLowerCharacterPeriod
/// * RomanUpperCharacterParenBoth
/// * RomanUpperCharacterParenR
/// * RomanUpperCharacterPeriod
/// * ThaiAlphaParenthesisBoth
/// * ThaiAlphaParenthesisRight
/// * ThaiAlphaPeriod
/// * ThaiNumberParenthesisBoth
/// * ThaiNumberParenthesisRight
/// * ThaiNumberPeriod
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum TextAutoNumberSchemeValues {
    AlphaLowerCharacterParenBoth,
    AlphaLowerCharacterParenR,
    AlphaLowerCharacterPeriod,
    AlphaUpperCharacterParenBoth,
    AlphaUpperCharacterParenR,
    AlphaUpperCharacterPeriod,
    Arabic1Minus,
    Arabic2Minus,
    ArabicDoubleBytePeriod,
    ArabicDoubleBytePlain,
    ArabicParenBoth,
    ArabicParenR,
    ArabicPeriod,
    ArabicPlain,
    CircleNumberDoubleBytePlain,
    CircleNumberWingdingsBlackPlain,
    CircleNumberWingdingsWhitePlain,
    EastAsianJapaneseDoubleBytePeriod,
    EastAsianJapaneseKoreanPeriod,
    EastAsianJapaneseKoreanPlain,
    EastAsianSimplifiedChinesePeriod,
    EastAsianSimplifiedChinesePlain,
    EastAsianTraditionalChinesePeriod,
    EastAsianTraditionalChinesePlain,
    Hebrew2Minus,
    HindiAlpha1Period,
    HindiAlphaPeriod,
    HindiNumberParenthesisRight,
    HindiNumPeriod,
    RomanLowerCharacterParenBoth,
    RomanLowerCharacterParenR,
    RomanLowerCharacterPeriod,
    RomanUpperCharacterParenBoth,
    RomanUpperCharacterParenR,
    RomanUpperCharacterPeriod,
    ThaiAlphaParenthesisBoth,
    ThaiAlphaParenthesisRight,
    ThaiAlphaPeriod,
    ThaiNumberParenthesisBoth,
    ThaiNumberParenthesisRight,
    ThaiNumberPeriod,
}

impl TextAutoNumberSchemeValues {
    pub(crate) fn default() -> Self {
        Self::ArabicPeriod
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else { return Self::default() };
        return match s.as_ref() {
            "alphaLcParenBoth" => Self::AlphaLowerCharacterParenBoth,
            "alphaLcParenR" => Self::AlphaLowerCharacterParenR,
            "alphaLcPeriod" => Self::AlphaLowerCharacterPeriod,
            "alphaUcParenBoth" => Self::AlphaUpperCharacterParenBoth,
            "alphaUcParenR" => Self::AlphaUpperCharacterParenR,
            "alphaUcPeriod" => Self::AlphaUpperCharacterPeriod,
            "arabic1Minus" => Self::Arabic1Minus,
            "arabic2Minus" => Self::Arabic2Minus,
            "arabicDbPeriod" => Self::ArabicDoubleBytePeriod,
            "arabicDbPlain" => Self::ArabicDoubleBytePlain,
            "arabicParenBoth" => Self::ArabicParenBoth,
            "arabicParenR" => Self::ArabicParenR,
            "arabicPeriod" => Self::ArabicPeriod,
            "arabicPlain" => Self::ArabicPlain,
            "circleNumDbPlain" => Self::CircleNumberDoubleBytePlain,
            "circleNumWdBlackPlain" => Self::CircleNumberWingdingsBlackPlain,
            "circleNumWdWhitePlain" => Self::CircleNumberWingdingsWhitePlain,
            "ea1JpnChsDbPeriod" => Self::EastAsianJapaneseDoubleBytePeriod,
            "ea1JpnKorPeriod" => Self::EastAsianJapaneseKoreanPeriod,
            "ea1JpnKorPlain" => Self::EastAsianJapaneseKoreanPlain,
            "ea1ChsPeriod" => Self::EastAsianSimplifiedChinesePeriod,
            "ea1ChsPlain" => Self::EastAsianSimplifiedChinesePlain,
            "ea1ChtPeriod" => Self::EastAsianTraditionalChinesePeriod,
            "ea1ChtPlain" => Self::EastAsianTraditionalChinesePlain,
            "hebrew2Minus" => Self::Hebrew2Minus,
            "hindiAlpha1Period" => Self::HindiAlpha1Period,
            "hindiAlphaPeriod" => Self::HindiAlphaPeriod,
            "hindiNumParenR" => Self::HindiNumberParenthesisRight,
            "hindiNumPeriod" => Self::HindiNumPeriod,
            "romanLcParenBoth" => Self::RomanLowerCharacterParenBoth,
            "romanLcParenR" => Self::RomanLowerCharacterParenR,
            "romanLcPeriod" => Self::RomanLowerCharacterPeriod,
            "romanUcParenBoth" => Self::RomanUpperCharacterParenBoth,
            "romanUcParenR" => Self::RomanUpperCharacterParenR,
            "romanUcPeriod" => Self::RomanUpperCharacterPeriod,
            "thaiAlphaParenBoth" => Self::ThaiAlphaParenthesisBoth,
            "thaiAlphaParenR" => Self::ThaiAlphaParenthesisRight,
            "thaiAlphaPeriod" => Self::ThaiAlphaPeriod,
            "thaiNumParenBoth" => Self::ThaiNumberParenthesisBoth,
            "thaiNumParenR" => Self::ThaiNumberParenthesisRight,
            "thaiNumPeriod" => Self::ThaiNumberPeriod,
            _ => Self::default(),
        };
    }
}
