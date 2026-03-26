namespace ExcelEngine.Models;

public enum SymbologyFamily
{
    Unknown,
    Linear1D,
    DataMatrix,
    GS1DataMatrix,
    QRCode,
    GS1QRCode,
    RectangularDataMatrix,
    DMRE,
    DotCode,
}

public enum SymbologyGroup
{
    Universal,
    Linear1D,
    TwoDCommon,
    TwoDDataMatrix,
    TwoDDataMatrixQuadrant,
    TwoDQR,
    TwoDRectangular,
    VendorPartTracking,
    MilStd,
}

public enum GradeLetterValue
{
    NotApplicable,
    A,
    B,
    C,
    D,
    F,
}

public enum OverallPassFail
{
    Pass,
    Fail,
    NotApplicable,
}

public enum OutputFormat
{
    Xlsx,
    Xls,
}

public enum ImagePolarity
{
    BlackOnWhite,
    WhiteOnBlack,
    Unknown,
}
