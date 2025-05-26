#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetpatternvalues?view=openxml-3.0.1
///
/// * Cross
/// * DarkDownwardDiagonal
/// * DarkHorizontal
/// * DarkUpwardDiagonal
/// * DarkVertical
/// * DashedDownwardDiagonal
/// * DashedHorizontal
/// * DashedUpwardDiagonal
/// * DashedVertical
/// * DiagonalBrick
/// * DiagonalCross
/// * Divot
/// * DotGrid
/// * DottedDiamond
/// * DownwardDiagonal
/// * Horizontal
/// * HorizontalBrick
/// * LargeCheck
/// * LargeConfetti
/// * LargeGrid
/// * LightDownwardDiagonal
/// * LightHorizontal
/// * LightUpwardDiagonal
/// * LightVertical
/// * NarrowHorizontal
/// * NarrowVertical
/// * OpenDiamond
/// * Percent10
/// * Percent20
/// * Percent25
/// * Percent30
/// * Percent40
/// * Percent5
/// * Percent50
/// * Percent60
/// * Percent70
/// * Percent75
/// * Percent80
/// * Percent90
/// * Plaid
/// * Shingle
/// * SmallCheck
/// * SmallConfetti
/// * SmallGrid
/// * SolidDiamond
/// * Sphere
/// * Trellis
/// * UpwardDiagonal
/// * Vertical
/// * Wave
/// * Weave
/// * WideDownwardDiagonal
/// * WideUpwardDiagonal
/// * ZigZag
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum PresetPatternValues {
    Cross,
    DarkDownwardDiagonal,
    DarkHorizontal,
    DarkUpwardDiagonal,
    DarkVertical,
    DashedDownwardDiagonal,
    DashedHorizontal,
    DashedUpwardDiagonal,
    DashedVertical,
    DiagonalBrick,
    DiagonalCross,
    Divot,
    DotGrid,
    DottedDiamond,
    DownwardDiagonal,
    Horizontal,
    HorizontalBrick,
    LargeCheck,
    LargeConfetti,
    LargeGrid,
    LightDownwardDiagonal,
    LightHorizontal,
    LightUpwardDiagonal,
    LightVertical,
    NarrowHorizontal,
    NarrowVertical,
    OpenDiamond,
    Percent10,
    Percent20,
    Percent25,
    Percent30,
    Percent40,
    Percent5,
    Percent50,
    Percent60,
    Percent70,
    Percent75,
    Percent80,
    Percent90,
    Plaid,
    Shingle,
    SmallCheck,
    SmallConfetti,
    SmallGrid,
    SolidDiamond,
    Sphere,
    Trellis,
    UpwardDiagonal,
    Vertical,
    Wave,
    Weave,
    WideDownwardDiagonal,
    WideUpwardDiagonal,
    ZigZag,
}

impl PresetPatternValues {
    pub(crate) fn from_string(s: Option<String>) -> Option<Self> {
        let Some(s) = s else { return None };
        return Some(match s.as_ref() {
            "cross" => Self::Cross,
            "dkDnDiag" => Self::DarkDownwardDiagonal,
            "dkHorz" => Self::DarkHorizontal,
            "dkUpDiag" => Self::DarkUpwardDiagonal,
            "dkVert" => Self::DarkVertical,
            "dashDnDiag" => Self::DashedDownwardDiagonal,
            "dashHorz" => Self::DashedHorizontal,
            "dashUpDiag" => Self::DashedUpwardDiagonal,
            "dashVert" => Self::DashedVertical,
            "diagBrick" => Self::DiagonalBrick,
            "diagCross" => Self::DiagonalCross,
            "divot" => Self::Divot,
            "dotGrid" => Self::DotGrid,
            "dotDmnd" => Self::DottedDiamond,
            "dnDiag" => Self::DownwardDiagonal,
            "horz" => Self::Horizontal,
            "horzBrick" => Self::HorizontalBrick,
            "lgCheck" => Self::LargeCheck,
            "lgConfetti" => Self::LargeConfetti,
            "lgGrid" => Self::LargeGrid,
            "ltDnDiag" => Self::LightDownwardDiagonal,
            "ltHorz" => Self::LightHorizontal,
            "ltUpDiag" => Self::LightUpwardDiagonal,
            "ltVert" => Self::LightVertical,
            "narHorz" => Self::NarrowHorizontal,
            "narVert" => Self::NarrowVertical,
            "openDmnd" => Self::OpenDiamond,
            "pct10" => Self::Percent10,
            "pct20" => Self::Percent20,
            "pct25" => Self::Percent25,
            "pct30" => Self::Percent30,
            "pct40" => Self::Percent40,
            "pct5" => Self::Percent5,
            "pct50" => Self::Percent50,
            "pct60" => Self::Percent60,
            "pct70" => Self::Percent70,
            "pct75" => Self::Percent75,
            "pct80" => Self::Percent80,
            "pct90" => Self::Percent90,
            "plaid" => Self::Plaid,
            "shingle" => Self::Shingle,
            "smCheck" => Self::SmallCheck,
            "smConfetti" => Self::SmallConfetti,
            "smGrid" => Self::SmallGrid,
            "solidDmnd" => Self::SolidDiamond,
            "sphere" => Self::Sphere,
            "trellis" => Self::Trellis,
            "upDiag" => Self::UpwardDiagonal,
            "vert" => Self::Vertical,
            "wave" => Self::Wave,
            "weave" => Self::Weave,
            "wdDnDiag" => Self::WideDownwardDiagonal,
            "wdUpDiag" => Self::WideUpwardDiagonal,
            "zigZag" => Self::ZigZag,
            _ => return None,
        });
    }
}
