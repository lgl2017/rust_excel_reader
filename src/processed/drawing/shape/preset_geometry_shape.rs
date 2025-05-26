#[cfg(feature = "serde")]
use serde::Serialize;

/// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.shapetypevalues?view=openxml-3.0.1
///
/// * AccentBorderCallout1
/// * AccentBorderCallout2
/// * AccentBorderCallout3
/// * AccentCallout1
/// * AccentCallout2
/// * AccentCallout3
/// * ActionButtonBackPrevious
/// * ActionButtonBeginning
/// * ActionButtonBlank
/// * ActionButtonDocument
/// * ActionButtonEnd
/// * ActionButtonForwardNext
/// * ActionButtonHelp
/// * ActionButtonHome
/// * ActionButtonInformation
/// * ActionButtonMovie
/// * ActionButtonReturn
/// * ActionButtonSound
/// * Arc
/// * BentArrow
/// * BentConnector2
/// * BentConnector3
/// * BentConnector4
/// * BentConnector5
/// * BentUpArrow
/// * Bevel
/// * BlockArc
/// * BorderCallout1
/// * BorderCallout2
/// * BorderCallout3
/// * BracePair
/// * BracketPair
/// * Callout1
/// * Callout2
/// * Callout3
/// * Can
/// * ChartPlus
/// * ChartStar
/// * ChartX
/// * Chevron
/// * Chord
/// * CircularArrow
/// * Cloud
/// * CloudCallout
/// * Corner
/// * CornerTabs
/// * Cube
/// * CurvedConnector2
/// * CurvedConnector3
/// * CurvedConnector4
/// * CurvedConnector5
/// * CurvedDownArrow
/// * CurvedLeftArrow
/// * CurvedRightArrow
/// * CurvedUpArrow
/// * Decagon
/// * DiagonalStripe
/// * Diamond
/// * Dodecagon
/// * Donut
/// * DoubleWave
/// * DownArrow
/// * DownArrowCallout
/// * Ellipse
/// * EllipseRibbon
/// * EllipseRibbon2
/// * FlowChartAlternateProcess
/// * FlowChartCollate
/// * FlowChartConnector
/// * FlowChartDecision
/// * FlowChartDelay
/// * FlowChartDisplay
/// * FlowChartDocument
/// * FlowChartExtract
/// * FlowChartInputOutput
/// * FlowChartInternalStorage
/// * FlowChartMagneticDisk
/// * FlowChartMagneticDrum
/// * FlowChartMagneticTape
/// * FlowChartManualInput
/// * FlowChartManualOperation
/// * FlowChartMerge
/// * FlowChartMultidocument
/// * FlowChartOfflineStorage
/// * FlowChartOffpageConnector
/// * FlowChartOnlineStorage
/// * FlowChartOr
/// * FlowChartPredefinedProcess
/// * FlowChartPreparation
/// * FlowChartProcess
/// * FlowChartPunchedCard
/// * FlowChartPunchedTape
/// * FlowChartSort
/// * FlowChartSummingJunction
/// * FlowChartTerminator
/// * FoldedCorner
/// * Frame
/// * Funnel
/// * Gear6
/// * Gear9
/// * HalfFrame
/// * Heart
/// * Heptagon
/// * Hexagon
/// * HomePlate
/// * HorizontalScroll
/// * IrregularSeal1
/// * IrregularSeal2
/// * LeftArrow
/// * LeftArrowCallout
/// * LeftBrace
/// * LeftBracket
/// * LeftCircularArrow
/// * LeftRightArrow
/// * LeftRightArrowCallout
/// * LeftRightCircularArrow
/// * LeftRightRibbon
/// * LeftRightUpArrow
/// * LeftUpArrow
/// * LightningBolt
/// * Line
/// * LineInverse
/// * MathDivide
/// * MathEqual
/// * MathMinus
/// * MathMultiply
/// * MathNotEqual
/// * MathPlus
/// * Moon
/// * NonIsoscelesTrapezoid
/// * NoSmoking
/// * NotchedRightArrow
/// * Octagon
/// * Parallelogram
/// * Pentagon
/// * Pie
/// * PieWedge
/// * Plaque
/// * PlaqueTabs
/// * Plus
/// * QuadArrow
/// * QuadArrowCallout
/// * Rectangle
/// * Ribbon
/// * Ribbon2
/// * RightArrow
/// * RightArrowCallout
/// * RightBrace
/// * RightBracket
/// * RightTriangle
/// * Round1Rectangle
/// * Round2DiagonalRectangle
/// * Round2SameRectangle
/// * RoundRectangle
/// * SmileyFace
/// * Snip1Rectangle
/// * Snip2DiagonalRectangle
/// * Snip2SameRectangle
/// * SnipRoundRectangle
/// * SquareTabs
/// * Star10
/// * Star12
/// * Star16
/// * Star24
/// * Star32
/// * Star4
/// * Star5
/// * Star6
/// * Star7
/// * Star8
/// * StraightConnector1
/// * StripedRightArrow
/// * Sun
/// * SwooshArrow
/// * Teardrop
/// * Trapezoid
/// * Triangle
/// * UpArrow
/// * UpArrowCallout
/// * UpDownArrow
/// * UpDownArrowCallout
/// * UTurnArrow
/// * VerticalScroll
/// * Wave
/// * WedgeEllipseCallout
/// * WedgeRectangleCallout
/// * WedgeRoundRectangleCallout
#[derive(Debug, Clone, PartialEq)]
#[cfg_attr(feature = "serde", derive(Serialize))]
#[cfg_attr(feature = "serde", serde(rename_all = "snake_case"))]
pub enum PresetShapeTypeValues {
    AccentBorderCallout1,
    AccentBorderCallout2,
    AccentBorderCallout3,
    AccentCallout1,
    AccentCallout2,
    AccentCallout3,
    ActionButtonBackPrevious,
    ActionButtonBeginning,
    ActionButtonBlank,
    ActionButtonDocument,
    ActionButtonEnd,
    ActionButtonForwardNext,
    ActionButtonHelp,
    ActionButtonHome,
    ActionButtonInformation,
    ActionButtonMovie,
    ActionButtonReturn,
    ActionButtonSound,
    Arc,
    BentArrow,
    BentConnector2,
    BentConnector3,
    BentConnector4,
    BentConnector5,
    BentUpArrow,
    Bevel,
    BlockArc,
    BorderCallout1,
    BorderCallout2,
    BorderCallout3,
    BracePair,
    BracketPair,
    Callout1,
    Callout2,
    Callout3,
    Can,
    ChartPlus,
    ChartStar,
    ChartX,
    Chevron,
    Chord,
    CircularArrow,
    Cloud,
    CloudCallout,
    Corner,
    CornerTabs,
    Cube,
    CurvedConnector2,
    CurvedConnector3,
    CurvedConnector4,
    CurvedConnector5,
    CurvedDownArrow,
    CurvedLeftArrow,
    CurvedRightArrow,
    CurvedUpArrow,
    Decagon,
    DiagonalStripe,
    Diamond,
    Dodecagon,
    Donut,
    DoubleWave,
    DownArrow,
    DownArrowCallout,
    Ellipse,
    EllipseRibbon,
    EllipseRibbon2,
    FlowChartAlternateProcess,
    FlowChartCollate,
    FlowChartConnector,
    FlowChartDecision,
    FlowChartDelay,
    FlowChartDisplay,
    FlowChartDocument,
    FlowChartExtract,
    FlowChartInputOutput,
    FlowChartInternalStorage,
    FlowChartMagneticDisk,
    FlowChartMagneticDrum,
    FlowChartMagneticTape,
    FlowChartManualInput,
    FlowChartManualOperation,
    FlowChartMerge,
    FlowChartMultidocument,
    FlowChartOfflineStorage,
    FlowChartOffpageConnector,
    FlowChartOnlineStorage,
    FlowChartOr,
    FlowChartPredefinedProcess,
    FlowChartPreparation,
    FlowChartProcess,
    FlowChartPunchedCard,
    FlowChartPunchedTape,
    FlowChartSort,
    FlowChartSummingJunction,
    FlowChartTerminator,
    FoldedCorner,
    Frame,
    Funnel,
    Gear6,
    Gear9,
    HalfFrame,
    Heart,
    Heptagon,
    Hexagon,
    HomePlate,
    HorizontalScroll,
    IrregularSeal1,
    IrregularSeal2,
    LeftArrow,
    LeftArrowCallout,
    LeftBrace,
    LeftBracket,
    LeftCircularArrow,
    LeftRightArrow,
    LeftRightArrowCallout,
    LeftRightCircularArrow,
    LeftRightRibbon,
    LeftRightUpArrow,
    LeftUpArrow,
    LightningBolt,
    Line,
    LineInverse,
    MathDivide,
    MathEqual,
    MathMinus,
    MathMultiply,
    MathNotEqual,
    MathPlus,
    Moon,
    NonIsoscelesTrapezoid,
    NoSmoking,
    NotchedRightArrow,
    Octagon,
    Parallelogram,
    Pentagon,
    Pie,
    PieWedge,
    Plaque,
    PlaqueTabs,
    Plus,
    QuadArrow,
    QuadArrowCallout,
    Rectangle,
    Ribbon,
    Ribbon2,
    RightArrow,
    RightArrowCallout,
    RightBrace,
    RightBracket,
    RightTriangle,
    Round1Rectangle,
    Round2DiagonalRectangle,
    Round2SameRectangle,
    RoundRectangle,
    SmileyFace,
    Snip1Rectangle,
    Snip2DiagonalRectangle,
    Snip2SameRectangle,
    SnipRoundRectangle,
    SquareTabs,
    Star10,
    Star12,
    Star16,
    Star24,
    Star32,
    Star4,
    Star5,
    Star6,
    Star7,
    Star8,
    StraightConnector1,
    StripedRightArrow,
    Sun,
    SwooshArrow,
    Teardrop,
    Trapezoid,
    Triangle,
    UpArrow,
    UpArrowCallout,
    UpDownArrow,
    UpDownArrowCallout,
    UTurnArrow,
    VerticalScroll,
    Wave,
    WedgeEllipseCallout,
    WedgeRectangleCallout,
    WedgeRoundRectangleCallout,
}

impl PresetShapeTypeValues {
    pub(crate) fn default() -> Self {
        Self::Rectangle
    }
    pub(crate) fn from_string(s: Option<String>) -> Self {
        let Some(s) = s else {
            return Self::default();
        };

        return match s.as_ref() {
            "accentBorderCallout1" => Self::AccentBorderCallout1,
            "accentBorderCallout2" => Self::AccentBorderCallout2,
            "accentBorderCallout3" => Self::AccentBorderCallout3,
            "accentCallout1" => Self::AccentCallout1,
            "accentCallout2" => Self::AccentCallout2,
            "accentCallout3" => Self::AccentCallout3,
            "actionButtonBackPrevious" => Self::ActionButtonBackPrevious,
            "actionButtonBeginning" => Self::ActionButtonBeginning,
            "actionButtonBlank" => Self::ActionButtonBlank,
            "actionButtonDocument" => Self::ActionButtonDocument,
            "actionButtonEnd" => Self::ActionButtonEnd,
            "actionButtonForwardNext" => Self::ActionButtonForwardNext,
            "actionButtonHelp" => Self::ActionButtonHelp,
            "actionButtonHome" => Self::ActionButtonHome,
            "actionButtonInformation" => Self::ActionButtonInformation,
            "actionButtonMovie" => Self::ActionButtonMovie,
            "actionButtonReturn" => Self::ActionButtonReturn,
            "actionButtonSound" => Self::ActionButtonSound,
            "arc" => Self::Arc,
            "bentArrow" => Self::BentArrow,
            "bentConnector2" => Self::BentConnector2,
            "bentConnector3" => Self::BentConnector3,
            "bentConnector4" => Self::BentConnector4,
            "bentConnector5" => Self::BentConnector5,
            "bentUpArrow" => Self::BentUpArrow,
            "bevel" => Self::Bevel,
            "blockArc" => Self::BlockArc,
            "borderCallout1" => Self::BorderCallout1,
            "borderCallout2" => Self::BorderCallout2,
            "borderCallout3" => Self::BorderCallout3,
            "bracePair" => Self::BracePair,
            "bracketPair" => Self::BracketPair,
            "callout1" => Self::Callout1,
            "callout2" => Self::Callout2,
            "callout3" => Self::Callout3,
            "can" => Self::Can,
            "chartPlus" => Self::ChartPlus,
            "chartStar" => Self::ChartStar,
            "chartX" => Self::ChartX,
            "chevron" => Self::Chevron,
            "chord" => Self::Chord,
            "circularArrow" => Self::CircularArrow,
            "cloud" => Self::Cloud,
            "cloudCallout" => Self::CloudCallout,
            "corner" => Self::Corner,
            "cornerTabs" => Self::CornerTabs,
            "cube" => Self::Cube,
            "curvedConnector2" => Self::CurvedConnector2,
            "curvedConnector3" => Self::CurvedConnector3,
            "curvedConnector4" => Self::CurvedConnector4,
            "curvedConnector5" => Self::CurvedConnector5,
            "curvedDownArrow" => Self::CurvedDownArrow,
            "curvedLeftArrow" => Self::CurvedLeftArrow,
            "curvedRightArrow" => Self::CurvedRightArrow,
            "curvedUpArrow" => Self::CurvedUpArrow,
            "decagon" => Self::Decagon,
            "diagStripe" => Self::DiagonalStripe,
            "diamond" => Self::Diamond,
            "dodecagon" => Self::Dodecagon,
            "donut" => Self::Donut,
            "doubleWave" => Self::DoubleWave,
            "downArrow" => Self::DownArrow,
            "downArrowCallout" => Self::DownArrowCallout,
            "ellipse" => Self::Ellipse,
            "ellipseRibbon" => Self::EllipseRibbon,
            "ellipseRibbon2" => Self::EllipseRibbon2,
            "flowChartAlternateProcess" => Self::FlowChartAlternateProcess,
            "flowChartCollate" => Self::FlowChartCollate,
            "flowChartConnector" => Self::FlowChartConnector,
            "flowChartDecision" => Self::FlowChartDecision,
            "flowChartDelay" => Self::FlowChartDelay,
            "flowChartDisplay" => Self::FlowChartDisplay,
            "flowChartDocument" => Self::FlowChartDocument,
            "flowChartExtract" => Self::FlowChartExtract,
            "flowChartInputOutput" => Self::FlowChartInputOutput,
            "flowChartInternalStorage" => Self::FlowChartInternalStorage,
            "flowChartMagneticDisk" => Self::FlowChartMagneticDisk,
            "flowChartMagneticDrum" => Self::FlowChartMagneticDrum,
            "flowChartMagneticTape" => Self::FlowChartMagneticTape,
            "flowChartManualInput" => Self::FlowChartManualInput,
            "flowChartManualOperation" => Self::FlowChartManualOperation,
            "flowChartMerge" => Self::FlowChartMerge,
            "flowChartMultidocument" => Self::FlowChartMultidocument,
            "flowChartOfflineStorage" => Self::FlowChartOfflineStorage,
            "flowChartOffpageConnector" => Self::FlowChartOffpageConnector,
            "flowChartOnlineStorage" => Self::FlowChartOnlineStorage,
            "flowChartOr" => Self::FlowChartOr,
            "flowChartPredefinedProcess" => Self::FlowChartPredefinedProcess,
            "flowChartPreparation" => Self::FlowChartPreparation,
            "flowChartProcess" => Self::FlowChartProcess,
            "flowChartPunchedCard" => Self::FlowChartPunchedCard,
            "flowChartPunchedTape" => Self::FlowChartPunchedTape,
            "flowChartSort" => Self::FlowChartSort,
            "flowChartSummingJunction" => Self::FlowChartSummingJunction,
            "flowChartTerminator" => Self::FlowChartTerminator,
            "foldedCorner" => Self::FoldedCorner,
            "frame" => Self::Frame,
            "funnel" => Self::Funnel,
            "gear6" => Self::Gear6,
            "gear9" => Self::Gear9,
            "halfFrame" => Self::HalfFrame,
            "heart" => Self::Heart,
            "heptagon" => Self::Heptagon,
            "hexagon" => Self::Hexagon,
            "homePlate" => Self::HomePlate,
            "horizontalScroll" => Self::HorizontalScroll,
            "irregularSeal1" => Self::IrregularSeal1,
            "irregularSeal2" => Self::IrregularSeal2,
            "leftArrow" => Self::LeftArrow,
            "leftArrowCallout" => Self::LeftArrowCallout,
            "leftBrace" => Self::LeftBrace,
            "leftBracket" => Self::LeftBracket,
            "leftCircularArrow" => Self::LeftCircularArrow,
            "leftRightArrow" => Self::LeftRightArrow,
            "leftRightArrowCallout" => Self::LeftRightArrowCallout,
            "leftRightCircularArrow" => Self::LeftRightCircularArrow,
            "leftRightRibbon" => Self::LeftRightRibbon,
            "leftRightUpArrow" => Self::LeftRightUpArrow,
            "leftUpArrow" => Self::LeftUpArrow,
            "lightningBolt" => Self::LightningBolt,
            "line" => Self::Line,
            "lineInv" => Self::LineInverse,
            "mathDivide" => Self::MathDivide,
            "mathEqual" => Self::MathEqual,
            "mathMinus" => Self::MathMinus,
            "mathMultiply" => Self::MathMultiply,
            "mathNotEqual" => Self::MathNotEqual,
            "mathPlus" => Self::MathPlus,
            "moon" => Self::Moon,
            "nonIsoscelesTrapezoid" => Self::NonIsoscelesTrapezoid,
            "noSmoking" => Self::NoSmoking,
            "notchedRightArrow" => Self::NotchedRightArrow,
            "octagon" => Self::Octagon,
            "parallelogram" => Self::Parallelogram,
            "pentagon" => Self::Pentagon,
            "pie" => Self::Pie,
            "pieWedge" => Self::PieWedge,
            "plaque" => Self::Plaque,
            "plaqueTabs" => Self::PlaqueTabs,
            "plus" => Self::Plus,
            "quadArrow" => Self::QuadArrow,
            "quadArrowCallout" => Self::QuadArrowCallout,
            "rect" => Self::Rectangle,
            "ribbon" => Self::Ribbon,
            "ribbon2" => Self::Ribbon2,
            "rightArrow" => Self::RightArrow,
            "rightArrowCallout" => Self::RightArrowCallout,
            "rightBrace" => Self::RightBrace,
            "rightBracket" => Self::RightBracket,
            "rtTriangle" => Self::RightTriangle,
            "round1Rect" => Self::Round1Rectangle,
            "round2DiagRect" => Self::Round2DiagonalRectangle,
            "round2SameRect" => Self::Round2SameRectangle,
            "roundRect" => Self::RoundRectangle,
            "smileyFace" => Self::SmileyFace,
            "snip1Rect" => Self::Snip1Rectangle,
            "snip2DiagRect" => Self::Snip2DiagonalRectangle,
            "snip2SameRect" => Self::Snip2SameRectangle,
            "snipRoundRect" => Self::SnipRoundRectangle,
            "squareTabs" => Self::SquareTabs,
            "star10" => Self::Star10,
            "star12" => Self::Star12,
            "star16" => Self::Star16,
            "star24" => Self::Star24,
            "star32" => Self::Star32,
            "star4" => Self::Star4,
            "star5" => Self::Star5,
            "star6" => Self::Star6,
            "star7" => Self::Star7,
            "star8" => Self::Star8,
            "straightConnector1" => Self::StraightConnector1,
            "stripedRightArrow" => Self::StripedRightArrow,
            "sun" => Self::Sun,
            "swooshArrow" => Self::SwooshArrow,
            "teardrop" => Self::Teardrop,
            "trapezoid" => Self::Trapezoid,
            "triangle" => Self::Triangle,
            "upArrow" => Self::UpArrow,
            "upArrowCallout" => Self::UpArrowCallout,
            "upDownArrow" => Self::UpDownArrow,
            "upDownArrowCallout" => Self::UpDownArrowCallout,
            "uturnArrow" => Self::UTurnArrow,
            "verticalScroll" => Self::VerticalScroll,
            "wave" => Self::Wave,
            "wedgeEllipseCallout" => Self::WedgeEllipseCallout,
            "wedgeRectCallout" => Self::WedgeRectangleCallout,
            "wedgeRoundRectCallout" => Self::WedgeRoundRectangleCallout,
            _ => Self::default(),
        };
    }
}
