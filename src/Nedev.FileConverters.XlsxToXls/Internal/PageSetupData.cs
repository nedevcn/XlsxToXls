namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Represents page setup data for BIFF8 format conversion.
/// Contains all information needed for printing and page layout.
/// </summary>
public sealed class PageSetupData
{
    /// <summary>Gets or sets the paper size. Default is Letter (1).</summary>
    public PaperSize PaperSize { get; set; } = PaperSize.Letter;

    /// <summary>Gets or sets the page orientation. Default is Portrait.</summary>
    public PageOrientation Orientation { get; set; } = PageOrientation.Portrait;

    /// <summary>Gets or sets the page scaling factor (10-400). Null means fit to pages.</summary>
    public int? Scale { get; set; }

    /// <summary>Gets or sets the number of pages wide to fit to.</summary>
    public int? FitToWidth { get; set; }

    /// <summary>Gets or sets the number of pages tall to fit to.</summary>
    public int? FitToHeight { get; set; }

    /// <summary>Gets or sets the page margins in inches.</summary>
    public PageMargins Margins { get; set; } = new();

    /// <summary>Gets or sets the print area cell ranges.</summary>
    public List<CellRange> PrintArea { get; set; } = [];

    /// <summary>Gets or sets the print title rows (repeated on each page).</summary>
    public CellRange? PrintTitleRows { get; set; }

    /// <summary>Gets or sets the print title columns (repeated on each page).</summary>
    public CellRange? PrintTitleColumns { get; set; }

    /// <summary>Gets or sets the header text.</summary>
    public string? Header { get; set; }

    /// <summary>Gets or sets the footer text.</summary>
    public string? Footer { get; set; }

    /// <summary>Gets or sets the header margin in inches. Default is 0.5.</summary>
    public double HeaderMargin { get; set; } = 0.5;

    /// <summary>Gets or sets the footer margin in inches. Default is 0.5.</summary>
    public double FooterMargin { get; set; } = 0.5;

    /// <summary>Gets or sets the first page number. Default is 1.</summary>
    public int FirstPageNumber { get; set; } = 1;

    /// <summary>Gets or sets the print quality in DPI. Default is 600.</summary>
    public int PrintQuality { get; set; } = 600;

    /// <summary>Gets or sets the starting page number. Null means auto.</summary>
    public int? StartPageNumber { get; set; }

    /// <summary>Gets or sets whether to center horizontally on page.</summary>
    public bool CenterHorizontally { get; set; }

    /// <summary>Gets or sets whether to center vertically on page.</summary>
    public bool CenterVertically { get; set; }

    /// <summary>Gets or sets the number of copies to print. Default is 1.</summary>
    public int Copies { get; set; } = 1;

    /// <summary>Gets or sets whether to print gridlines. Default is true.</summary>
    public bool PrintGridlines { get; set; } = true;

    /// <summary>Gets or sets whether to print row and column headings. Default is false.</summary>
    public bool PrintHeadings { get; set; }

    /// <summary>Gets or sets whether to print in black and white. Default is false.</summary>
    public bool BlackAndWhite { get; set; }

    /// <summary>Gets or sets whether to print comments. Default is None.</summary>
    public PrintComments PrintComments { get; set; } = PrintComments.None;

    /// <summary>Gets or sets the page order. Default is DownThenOver.</summary>
    public PageOrder PageOrder { get; set; } = PageOrder.DownThenOver;

    /// <summary>Gets or sets the cell errors print option. Default is Displayed.</summary>
    public CellErrorPrint CellErrors { get; set; } = CellErrorPrint.Displayed;

    /// <summary>Gets or sets whether to use draft quality. Default is false.</summary>
    public bool DraftQuality { get; set; }
}

/// <summary>Paper sizes supported by Excel.</summary>
public enum PaperSize : ushort
{
    /// <summary>Letter (8.5" x 11").</summary>
    Letter = 1,

    /// <summary>Letter Small (8.5" x 11").</summary>
    LetterSmall = 2,

    /// <summary>Tabloid (11" x 17").</summary>
    Tabloid = 3,

    /// <summary>Ledger (17" x 11").</summary>
    Ledger = 4,

    /// <summary>Legal (8.5" x 14").</summary>
    Legal = 5,

    /// <summary>Statement (5.5" x 8.5").</summary>
    Statement = 6,

    /// <summary>Executive (7.25" x 10.5").</summary>
    Executive = 7,

    /// <summary>A3 (297mm x 420mm).</summary>
    A3 = 8,

    /// <summary>A4 (210mm x 297mm).</summary>
    A4 = 9,

    /// <summary>A4 Small (210mm x 297mm).</summary>
    A4Small = 10,

    /// <summary>A5 (148mm x 210mm).</summary>
    A5 = 11,

    /// <summary>B4 (250mm x 353mm).</summary>
    B4 = 12,

    /// <summary>B5 (176mm x 250mm).</summary>
    B5 = 13,

    /// <summary>Folio (8.5" x 13").</summary>
    Folio = 14,

    /// <summary>Quarto (215mm x 275mm).</summary>
    Quarto = 15,

    /// <summary>Standard (10" x 14").</summary>
    Standard10x14 = 16,

    /// <summary>Standard (11" x 17").</summary>
    Standard11x17 = 17,

    /// <summary>Note (8.5" x 11").</summary>
    Note = 18,

    /// <summary>Envelope #9 (3.875" x 8.875").</summary>
    Envelope9 = 19,

    /// <summary>Envelope #10 (4.125" x 9.5").</summary>
    Envelope10 = 20,

    /// <summary>Envelope #11 (4.5" x 10.375").</summary>
    Envelope11 = 21,

    /// <summary>Envelope #12 (4.75" x 11").</summary>
    Envelope12 = 22,

    /// <summary>Envelope #14 (5" x 11.5").</summary>
    Envelope14 = 23,

    /// <summary>C size sheet.</summary>
    C = 24,

    /// <summary>D size sheet.</summary>
    D = 25,

    /// <summary>E size sheet.</summary>
    E = 26,

    /// <summary>Envelope DL (110mm x 220mm).</summary>
    EnvelopeDL = 27,

    /// <summary>Envelope C5 (162mm x 229mm).</summary>
    EnvelopeC5 = 28,

    /// <summary>Envelope C3 (324mm x 458mm).</summary>
    EnvelopeC3 = 29,

    /// <summary>Envelope C4 (229mm x 324mm).</summary>
    EnvelopeC4 = 30,

    /// <summary>Envelope C6 (114mm x 162mm).</summary>
    EnvelopeC6 = 31,

    /// <summary>Envelope C65 (114mm x 229mm).</summary>
    EnvelopeC65 = 32,

    /// <summary>Envelope B4 (250mm x 353mm).</summary>
    EnvelopeB4 = 33,

    /// <summary>Envelope B5 (176mm x 250mm).</summary>
    EnvelopeB5 = 34,

    /// <summary>Envelope B6 (125mm x 176mm).</summary>
    EnvelopeB6 = 35,

    /// <summary>Envelope (110mm x 230mm).</summary>
    Envelope = 36,

    /// <summary>Envelope Monarch (3.875" x 7.5").</summary>
    EnvelopeMonarch = 37,

    /// <summary>6.75 Envelope (3.625" x 6.5").</summary>
    Envelope67 = 38,

    /// <summary>US Standard Fanfold (14.875" x 11").</summary>
    USStandardFanfold = 39,

    /// <summary>German Standard Fanfold (8.5" x 12").</summary>
    GermanStandardFanfold = 40,

    /// <summary>German Legal Fanfold (8.5" x 13").</summary>
    GermanLegalFanfold = 41,

    /// <summary>B4 (ISO) (250mm x 353mm).</summary>
    B4ISO = 42,

    /// <summary>Japanese Postcard (100mm x 148mm).</summary>
    JapanesePostcard = 43,

    /// <summary>9 x 11 inch.</summary>
    Size9x11 = 44,

    /// <summary>10 x 11 inch.</summary>
    Size10x11 = 45,

    /// <summary>15 x 11 inch.</summary>
    Size15x11 = 46,

    /// <summary>Envelope Invite (220mm x 220mm).</summary>
    EnvelopeInvite = 47,

    /// <summary>Letter Extra (9.275" x 12").</summary>
    LetterExtra = 50,

    /// <summary>Legal Extra (9.275" x 15").</summary>
    LegalExtra = 51,

    /// <summary>Tabloid Extra (11.69" x 18").</summary>
    TabloidExtra = 52,

    /// <summary>A4 Extra (235mm x 322mm).</summary>
    A4Extra = 53,

    /// <summary>Letter Transverse (8.275" x 11").</summary>
    LetterTransverse = 54,

    /// <summary>A4 Transverse (210mm x 297mm).</summary>
    A4Transverse = 55,

    /// <summary>Letter Extra Transverse (9.275" x 12").</summary>
    LetterExtraTransverse = 56,

    /// <summary>SuperA (227mm x 356mm).</summary>
    SuperA = 57,

    /// <summary>SuperB (305mm x 487mm).</summary>
    SuperB = 58,

    /// <summary>US Letter Plus (8.5" x 12.69").</summary>
    USLetterPlus = 59,

    /// <summary>A4 Plus (210mm x 330mm).</summary>
    A4Plus = 60,

    /// <summary>A5 Transverse (148mm x 210mm).</summary>
    A5Transverse = 61,

    /// <summary>B5 Transverse (182mm x 257mm).</summary>
    B5Transverse = 62,

    /// <summary>A3 Extra (322mm x 445mm).</summary>
    A3Extra = 63,

    /// <summary>A5 Extra (174mm x 235mm).</summary>
    A5Extra = 64,

    /// <summary>B5 Extra (201mm x 276mm).</summary>
    B5Extra = 65,

    /// <summary>A2 (420mm x 594mm).</summary>
    A2 = 66,

    /// <summary>A3 Transverse (297mm x 420mm).</summary>
    A3Transverse = 67,

    /// <summary>A3 Extra Transverse (322mm x 445mm).</summary>
    A3ExtraTransverse = 68,

    /// <summary>Japanese Double Postcard (200mm x 148mm).</summary>
    JapaneseDoublePostcard = 69,

    /// <summary>A6 (105mm x 148mm).</summary>
    A6 = 70,

    /// <summary>Japanese Envelope Kaku #2.</summary>
    JapaneseEnvelopeKaku2 = 71,

    /// <summary>Japanese Envelope Kaku #3.</summary>
    JapaneseEnvelopeKaku3 = 72,

    /// <summary>Japanese Envelope Chou #3.</summary>
    JapaneseEnvelopeChou3 = 73,

    /// <summary>Japanese Envelope Chou #4.</summary>
    JapaneseEnvelopeChou4 = 74,

    /// <summary>Letter Rotated (11" x 8.5").</summary>
    LetterRotated = 75,

    /// <summary>A3 Rotated (420mm x 297mm).</summary>
    A3Rotated = 76,

    /// <summary>A4 Rotated (297mm x 210mm).</summary>
    A4Rotated = 77,

    /// <summary>A5 Rotated (210mm x 148mm).</summary>
    A5Rotated = 78,

    /// <summary>B4 Rotated (364mm x 257mm).</summary>
    B4Rotated = 79,

    /// <summary>B5 Rotated (257mm x 182mm).</summary>
    B5Rotated = 80,

    /// <summary>Japanese Postcard Rotated (148mm x 100mm).</summary>
    JapanesePostcardRotated = 81,

    /// <summary>Double Japanese Postcard Rotated (148mm x 200mm).</summary>
    DoubleJapanesePostcardRotated = 82,

    /// <summary>A6 Rotated (148mm x 105mm).</summary>
    A6Rotated = 83,

    /// <summary>Japanese Envelope Kaku #2 Rotated.</summary>
    JapaneseEnvelopeKaku2Rotated = 84,

    /// <summary>Japanese Envelope Kaku #3 Rotated.</summary>
    JapaneseEnvelopeKaku3Rotated = 85,

    /// <summary>Japanese Envelope Chou #3 Rotated.</summary>
    JapaneseEnvelopeChou3Rotated = 86,

    /// <summary>Japanese Envelope Chou #4 Rotated.</summary>
    JapaneseEnvelopeChou4Rotated = 87,

    /// <summary>B6 (125mm x 176mm).</summary>
    B6 = 88,

    /// <summary>B6 Rotated (176mm x 125mm).</summary>
    B6Rotated = 89,

    /// <summary>12 x 11 inch.</summary>
    Size12x11 = 90,

    /// <summary>Japanese Envelope You #4.</summary>
    JapaneseEnvelopeYou4 = 91,

    /// <summary>Japanese Envelope You #4 Rotated.</summary>
    JapaneseEnvelopeYou4Rotated = 92,

    /// <summary>PRC 16K (146mm x 215mm).</summary>
    PRC16K = 93,

    /// <summary>PRC 32K (97mm x 151mm).</summary>
    PRC32K = 94,

    /// <summary>PRC 32K(Big) (97mm x 151mm).</summary>
    PRC32KBig = 95,

    /// <summary>PRC Envelope #1 (102mm x 165mm).</summary>
    PRCEnvelope1 = 96,

    /// <summary>PRC Envelope #2 (102mm x 176mm).</summary>
    PRCEnvelope2 = 97,

    /// <summary>PRC Envelope #3 (125mm x 176mm).</summary>
    PRCEnvelope3 = 98,

    /// <summary>PRC Envelope #4 (110mm x 208mm).</summary>
    PRCEnvelope4 = 99,

    /// <summary>PRC Envelope #5 (110mm x 220mm).</summary>
    PRCEnvelope5 = 100,

    /// <summary>PRC Envelope #6 (120mm x 230mm).</summary>
    PRCEnvelope6 = 101,

    /// <summary>PRC Envelope #7 (160mm x 230mm).</summary>
    PRCEnvelope7 = 102,

    /// <summary>PRC Envelope #8 (120mm x 309mm).</summary>
    PRCEnvelope8 = 103,

    /// <summary>PRC Envelope #9 (229mm x 324mm).</summary>
    PRCEnvelope9 = 104,

    /// <summary>PRC Envelope #10 (324mm x 458mm).</summary>
    PRCEnvelope10 = 105,

    /// <summary>PRC 16K Rotated.</summary>
    PRC16KRotated = 106,

    /// <summary>PRC 32K Rotated.</summary>
    PRC32KRotated = 107,

    /// <summary>PRC 32K(Big) Rotated.</summary>
    PRC32KBigRotated = 108,

    /// <summary>PRC Envelope #1 Rotated (165mm x 102mm).</summary>
    PRCEnvelope1Rotated = 109,

    /// <summary>PRC Envelope #2 Rotated (176mm x 102mm).</summary>
    PRCEnvelope2Rotated = 110,

    /// <summary>PRC Envelope #3 Rotated (176mm x 125mm).</summary>
    PRCEnvelope3Rotated = 111,

    /// <summary>PRC Envelope #4 Rotated (208mm x 110mm).</summary>
    PRCEnvelope4Rotated = 112,

    /// <summary>PRC Envelope #5 Rotated (220mm x 110mm).</summary>
    PRCEnvelope5Rotated = 113,

    /// <summary>PRC Envelope #6 Rotated (230mm x 120mm).</summary>
    PRCEnvelope6Rotated = 114,

    /// <summary>PRC Envelope #7 Rotated (230mm x 160mm).</summary>
    PRCEnvelope7Rotated = 115,

    /// <summary>PRC Envelope #8 Rotated (309mm x 120mm).</summary>
    PRCEnvelope8Rotated = 116,

    /// <summary>PRC Envelope #9 Rotated (324mm x 229mm).</summary>
    PRCEnvelope9Rotated = 117,

    /// <summary>PRC Envelope #10 Rotated (458mm x 324mm).</summary>
    PRCEnvelope10Rotated = 118
}

/// <summary>Page orientation.</summary>
public enum PageOrientation : byte
{
    /// <summary>Portrait orientation.</summary>
    Portrait = 0,

    /// <summary>Landscape orientation.</summary>
    Landscape = 1
}

/// <summary>Page margins in inches.</summary>
public sealed class PageMargins
{
    /// <summary>Gets or sets the left margin. Default is 0.75 inches.</summary>
    public double Left { get; set; } = 0.75;

    /// <summary>Gets or sets the right margin. Default is 0.75 inches.</summary>
    public double Right { get; set; } = 0.75;

    /// <summary>Gets or sets the top margin. Default is 1.0 inches.</summary>
    public double Top { get; set; } = 1.0;

    /// <summary>Gets or sets the bottom margin. Default is 1.0 inches.</summary>
    public double Bottom { get; set; } = 1.0;

    /// <summary>Gets or sets the header margin. Default is 0.5 inches.</summary>
    public double Header { get; set; } = 0.5;

    /// <summary>Gets or sets the footer margin. Default is 0.5 inches.</summary>
    public double Footer { get; set; } = 0.5;
}

/// <summary>Page order for printing.</summary>
public enum PageOrder : byte
{
    /// <summary>Print down then across.</summary>
    DownThenOver = 0,

    /// <summary>Print across then down.</summary>
    OverThenDown = 1
}

/// <summary>Print comments options.</summary>
public enum PrintComments : byte
{
    /// <summary>Do not print comments.</summary>
    None = 0,

    /// <summary>Print comments at end of sheet.</summary>
    AtEnd = 1,

    /// <summary>Print comments as displayed.</summary>
    AsDisplayed = 2
}

/// <summary>Cell error print options.</summary>
public enum CellErrorPrint : byte
{
    /// <summary>Print errors as displayed.</summary>
    Displayed = 0,

    /// <summary>Print blank for errors.</summary>
    Blank = 1,

    /// <summary>Print -- for errors.</summary>
    DashDash = 2,

    /// <summary>Print #N/A for errors.</summary>
    NA = 3
}

/// <summary>
/// Header and footer formatting codes.
/// </summary>
public static class HeaderFooterCodes
{
    /// <summary>Left section.</summary>
    public const string Left = "&L";

    /// <summary>Center section.</summary>
    public const string Center = "&C";

    /// <summary>Right section.</summary>
    public const string Right = "&R";

    /// <summary>Page number.</summary>
    public const string PageNumber = "&P";

    /// <summary>Total pages.</summary>
    public const string TotalPages = "&N";

    /// <summary>Current date.</summary>
    public const string Date = "&D";

    /// <summary>Current time.</summary>
    public const string Time = "&T";

    /// <summary>File path.</summary>
    public const string FilePath = "&Z";

    /// <summary>File name.</summary>
    public const string FileName = "&F";

    /// <summary>Sheet name.</summary>
    public const string SheetName = "&A";

    /// <summary>Bold font.</summary>
    public const string Bold = "&B";

    /// <summary>Italic font.</summary>
    public const string Italic = "&I";

    /// <summary>Underline.</summary>
    public const string Underline = "&U";

    /// <summary>Strikethrough.</summary>
    public const string Strikethrough = "&S";

    /// <summary>Font size (e.g., &"10").</summary>
    public static string FontSize(int size) => $"&\"{size}\"";

    /// <summary>Font name (e.g., &"Arial").</summary>
    public static string FontName(string name) => $"&\"{name}\"";
}
