namespace Nedev.FileConverters.XlsxToXls.Internal;

internal sealed class StylesData
{
    private const int MinFonts = 4;

    public List<FontInfo> Fonts { get; } = [];

    public int FontOffset { get; private set; }

    public void EnsureMinFonts()
    {
        FontOffset = Math.Max(0, MinFonts - Fonts.Count);
        for (var i = 0; i < FontOffset; i++)
            Fonts.Insert(0, new FontInfo("Arial", 10, false, false, -1));
    }
    public List<NumFmtInfo> NumFmts { get; } = [];
    public List<CellXfInfo> CellXfs { get; } = [];

    public int GetBiffFontIndex(int xlsxFontId)
    {
        if (xlsxFontId < 0) return 0;
        return Math.Min(FontOffset + xlsxFontId, Fonts.Count - 1);
    }

    public int GetBiffFormatIndex(int xlsxNumFmtId)
    {
        if (xlsxNumFmtId < 0) return 0;
        var idx = NumFmts.FindIndex(n => n.NumFmtId == xlsxNumFmtId);
        if (idx >= 0) return 164 + idx;
        if (xlsxNumFmtId < 164) return xlsxNumFmtId;
        return 0;
    }

    public int GetBiffXfIndex(int xlsxStyleIndex)
    {
        if (xlsxStyleIndex < 0 || xlsxStyleIndex >= CellXfs.Count) return 15;
        return 15 + xlsxStyleIndex;
    }
}

internal record struct FontInfo(string Name, double Height, bool Bold, bool Italic, int ColorIndex);

internal record struct NumFmtInfo(int NumFmtId, string FormatCode);

internal record struct CellXfInfo(
    int NumFmtId,
    int FontId,
    int FillId,
    int BorderId,
    byte HorizontalAlign,
    byte VerticalAlign,
    bool WrapText,
    byte Indent,
    bool Locked,
    bool Hidden);
