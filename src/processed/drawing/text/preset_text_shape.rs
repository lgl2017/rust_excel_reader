#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.textshapevalues?view=openxml-3.0.1
///
/// * TextArchDown
/// * TextArchDownPour
/// * TextArchUp
/// * TextArchUpPour
/// * TextButton
/// * TextButtonPour
/// * TextCanDown
/// * TextCanUp
/// * TextCascadeDown
/// * TextCascadeUp
/// * TextChevron
/// * TextChevronInverted
/// * TextCircle
/// * TextCirclePour
/// * TextCurveDown
/// * TextCurveUp
/// * TextDeflate
/// * TextDeflateBottom
/// * TextDeflateInflate
/// * TextDeflateInflateDeflate
/// * TextDeflateTop
/// * TextDoubleWave1
/// * TextFadeDown
/// * TextFadeLeft
/// * TextFadeRight
/// * TextFadeUp
/// * TextInflate
/// * TextInflateBottom
/// * TextInflateTop
/// * TextNoShape
/// * TextPlain
/// * TextRingInside
/// * TextRingOutside
/// * TextSlantDown
/// * TextSlantUp
/// * TextStop
/// * TextTriangle
/// * TextTriangleInverted
/// * TextWave1
/// * TextWave2
/// * TextWave4
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum PresetTextShapeValues {
    TextArchDown,
    TextArchDownPour,
    TextArchUp,
    TextArchUpPour,
    TextButton,
    TextButtonPour,
    TextCanDown,
    TextCanUp,
    TextCascadeDown,
    TextCascadeUp,
    TextChevron,
    TextChevronInverted,
    TextCircle,
    TextCirclePour,
    TextCurveDown,
    TextCurveUp,
    TextDeflate,
    TextDeflateBottom,
    TextDeflateInflate,
    TextDeflateInflateDeflate,
    TextDeflateTop,
    TextDoubleWave1,
    TextFadeDown,
    TextFadeLeft,
    TextFadeRight,
    TextFadeUp,
    TextInflate,
    TextInflateBottom,
    TextInflateTop,
    TextNoShape,
    TextPlain,
    TextRingInside,
    TextRingOutside,
    TextSlantDown,
    TextSlantUp,
    TextStop,
    TextTriangle,
    TextTriangleInverted,
    TextWave1,
    TextWave2,
    TextWave4,
}

impl PresetTextShapeValues {
    pub(crate) fn default() -> Self {
        Self::TextPlain
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "textArchDown" => Self::TextArchDown,
            "textArchDownPour" => Self::TextArchDownPour,
            "textArchUp" => Self::TextArchUp,
            "textArchUpPour" => Self::TextArchUpPour,
            "textButton" => Self::TextButton,
            "textButtonPour" => Self::TextButtonPour,
            "textCanDown" => Self::TextCanDown,
            "textCanUp" => Self::TextCanUp,
            "textCascadeDown" => Self::TextCascadeDown,
            "textCascadeUp" => Self::TextCascadeUp,
            "textChevron" => Self::TextChevron,
            "textChevronInverted" => Self::TextChevronInverted,
            "textCircle" => Self::TextCircle,
            "textCirclePour" => Self::TextCirclePour,
            "textCurveDown" => Self::TextCurveDown,
            "textCurveUp" => Self::TextCurveUp,
            "textDeflate" => Self::TextDeflate,
            "textDeflateBottom" => Self::TextDeflateBottom,
            "textDeflateInflate" => Self::TextDeflateInflate,
            "textDeflateTop" => Self::TextDeflateTop,
            "textDoubleWave1" => Self::TextDoubleWave1,
            "textFadeDown" => Self::TextFadeDown,
            "textFadeLeft" => Self::TextFadeLeft,
            "textFadeRight" => Self::TextFadeRight,
            "textFadeUp" => Self::TextFadeUp,
            "textInflate" => Self::TextInflate,
            "textInflateBottom" => Self::TextInflateBottom,
            "textInflateTop" => Self::TextInflateTop,
            "textNoShape" => Self::TextNoShape,
            "textPlain" => Self::TextPlain,
            "textRingInside" => Self::TextRingInside,
            "textRingOutside" => Self::TextRingOutside,
            "textSlantDown" => Self::TextSlantDown,
            "textSlantUp" => Self::TextSlantUp,
            "textStop" => Self::TextStop,
            "textTriangle" => Self::TextTriangle,
            "textTriangleInverted" => Self::TextTriangleInverted,
            "textWave1" => Self::TextWave1,
            "textWave2" => Self::TextWave2,
            "textWave4" => Self::TextWave4,
            _ => Self::default(),
        };
    }
}
