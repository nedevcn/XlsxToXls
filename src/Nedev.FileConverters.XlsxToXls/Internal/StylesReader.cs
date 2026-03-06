using System.Globalization;
using System.IO.Compression;
using System.Xml;

namespace Nedev.FileConverters.XlsxToXls.Internal;

internal static class StylesReader
{
    private static readonly XmlReaderSettings Settings = new()
    {
        Async = false,
        CloseInput = false,
        IgnoreWhitespace = true,
        DtdProcessing = DtdProcessing.Prohibit
    };

    public static StylesData? Read(ZipArchive archive)
    {
        var entry = archive.GetEntry("xl/styles.xml");
        if (entry == null) return null;

        var styles = new StylesData();
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Settings);
        var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element) continue;
            if (reader.NamespaceURI != ns) continue;

            switch (reader.LocalName)
            {
                case "numFmt":
                    ReadNumFmt(reader, ns, styles);
                    break;
                case "font":
                    ReadFont(reader, ns, styles);
                    break;
                case "cellXf":
                    ReadCellXf(reader, ns, styles);
                    break;
            }
        }

        styles.EnsureMinFonts();
        if (styles.CellXfs.Count == 0)
            styles.CellXfs.Add(new CellXfInfo(0, 0, 0, 0, 0, 2, false, 0, true, false));
        return styles;
    }

    private static void ReadNumFmt(XmlReader reader, string ns, StylesData styles)
    {
        var id = ParseInt(reader.GetAttribute("numFmtId"));
        var code = reader.GetAttribute("formatCode") ?? "General";
        if (id >= 164)
            styles.NumFmts.Add(new NumFmtInfo(id, code));
    }

    private static void ReadFont(XmlReader reader, string ns, StylesData styles)
    {
        if (reader.IsEmptyElement) { styles.Fonts.Add(new FontInfo("Calibri", 11, false, false, -1)); return; }

        var name = "Calibri";
        var height = 11.0;
        var bold = false;
        var italic = false;
        var colorIndex = -1;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "font") break;
            if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != ns) continue;

            switch (reader.LocalName)
            {
                case "sz":
                    if (reader.Read() && reader.NodeType == XmlNodeType.Text && double.TryParse(reader.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var sz))
                        height = sz;
                    break;
                case "b":
                    bold = true;
                    break;
                case "i":
                    italic = true;
                    break;
                case "name":
                    var n = reader.GetAttribute("val");
                    if (!string.IsNullOrEmpty(n)) name = n;
                    break;
                case "color":
                    var theme = reader.GetAttribute("theme");
                    var rgb = reader.GetAttribute("rgb");
                    if (!string.IsNullOrEmpty(rgb)) colorIndex = ParseColorRgb(rgb);
                    break;
            }
        }

        styles.Fonts.Add(new FontInfo(name, height, bold, italic, colorIndex));
    }

    private static void ReadCellXf(XmlReader reader, string ns, StylesData styles)
    {
        var numFmtId = ParseInt(reader.GetAttribute("numFmtId"));
        var fontId = ParseInt(reader.GetAttribute("fontId"));
        var fillId = ParseInt(reader.GetAttribute("fillId"));
        var borderId = ParseInt(reader.GetAttribute("borderId"));
        if (numFmtId < 0) numFmtId = 0;
        if (fontId < 0) fontId = 0;

        byte horizontalAlign = 0;
        byte verticalAlign = 0;
        bool wrapText = false;
        byte indent = 0;
        bool locked = true;   // Excel 默认锁定
        bool hidden = false;

        if (!reader.IsEmptyElement)
        {
            var depth = reader.Depth;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "cellXf" && reader.NamespaceURI == ns && reader.Depth == depth)
                    break;
                if (reader.NodeType != XmlNodeType.Element || reader.NamespaceURI != ns) continue;

                if (reader.LocalName == "alignment")
                {
                    var h = reader.GetAttribute("horizontal");
                    var v = reader.GetAttribute("vertical");
                    var wrap = reader.GetAttribute("wrapText");
                    var ind = reader.GetAttribute("indent");

                    horizontalAlign = MapHorizontalAlignment(h);
                    verticalAlign = MapVerticalAlignment(v);
                    wrapText = ParseBool(wrap, defaultValue: false);
                    if (int.TryParse(ind, out var indVal) && indVal > 0)
                        indent = (byte)Math.Clamp(indVal, 0, 15);
                }
                else if (reader.LocalName == "protection")
                {
                    var lockedAttr = reader.GetAttribute("locked");
                    var hiddenAttr = reader.GetAttribute("hidden");
                    // locked 缺省通常为 true
                    locked = ParseBool(lockedAttr, defaultValue: true);
                    hidden = ParseBool(hiddenAttr, defaultValue: false);
                }
            }
        }

        styles.CellXfs.Add(new CellXfInfo(
            numFmtId,
            fontId,
            fillId,
            borderId,
            horizontalAlign,
            verticalAlign,
            wrapText,
            indent,
            locked,
            hidden));
    }

    private static int ParseInt(string? s)
    {
        if (string.IsNullOrEmpty(s) || !int.TryParse(s, out var v)) return -1;
        return v;
    }

    private static bool ParseBool(string? s, bool defaultValue)
    {
        if (string.IsNullOrEmpty(s)) return defaultValue;
        return s == "1" || s.Equals("true", StringComparison.OrdinalIgnoreCase);
    }

    private static byte MapHorizontalAlignment(string? value)
    {
        return value switch
        {
            "left" => 1,
            "center" => 2,
            "right" => 3,
            "fill" => 4,
            "justify" => 5,
            "centerContinuous" => 6,
            "distributed" => 7,
            _ => 0 // general
        };
    }

    private static byte MapVerticalAlignment(string? value)
    {
        return value switch
        {
            "top" => 0,
            "center" => 1,
            "bottom" => 2,
            "justify" => 3,
            "distributed" => 4,
            _ => 2 // default bottom
        };
    }

    private static int ParseColorRgb(string rgb)
    {
        if (rgb.Length >= 8 && rgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase))
            rgb = rgb[2..];
        if (rgb.Length != 6) return -1;
        return Convert.ToInt32(rgb, 16);
    }
}
