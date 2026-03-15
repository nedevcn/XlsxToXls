using System.IO.Compression;
using System.Xml;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// XLSX条件格式读取器 - 从Open XML格式读取条件格式数据
/// </summary>
internal static class ConditionalFormatReader
{
    private static readonly XmlReaderSettings Settings = new()
    {
        Async = false,
        CloseInput = false,
        IgnoreWhitespace = true,
        DtdProcessing = DtdProcessing.Prohibit
    };

    public static List<ConditionalFormatData> ReadConditionalFormats(ZipArchive archive, string worksheetPath, Action<string>? log = null)
    {
        var formats = new List<ConditionalFormatData>();
        try
        {
            var entry = archive.GetEntry(worksheetPath) ?? archive.GetEntry("xl/" + worksheetPath);
            if (entry == null)
            {
                log?.Invoke($"[ConditionalFormatReader] Worksheet not found: {worksheetPath}");
                return formats;
            }

            log?.Invoke($"[ConditionalFormatReader] Reading conditional formats from: {worksheetPath}");

            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, Settings);
            var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "conditionalFormatting" && reader.NamespaceURI == ns)
                {
                    var format = ReadConditionalFormatting(reader, ns, log);
                    if (format != null)
                    {
                        formats.Add(format);
                        log?.Invoke($"[ConditionalFormatReader] Found conditional format: {format.Type}");
                    }
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[ConditionalFormatReader] Error reading conditional formats: {ex.Message}");
        }

        return formats;
    }

    private static ConditionalFormatData? ReadConditionalFormatting(XmlReader reader, string ns, Action<string>? log)
    {
        var format = new ConditionalFormatData();

        // 读取范围
        var sqref = reader.GetAttribute("sqref");
        if (!string.IsNullOrEmpty(sqref))
        {
            format.Ranges = ParseRanges(sqref);
        }

        if (reader.IsEmptyElement) return format.Ranges.Count > 0 ? format : null;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                switch (reader.LocalName)
                {
                    case "cfRule":
                        ReadCfRule(reader, ns, format, log);
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }

        return format.Ranges.Count > 0 ? format : null;
    }

    private static void ReadCfRule(XmlReader reader, string ns, ConditionalFormatData format, Action<string>? log)
    {
        // 读取类型
        var type = reader.GetAttribute("type");
        format.Type = type switch
        {
            "cellIs" => ConditionalFormatType.CellIs,
            "containsText" => ConditionalFormatType.ContainsText,
            "notContainsText" => ConditionalFormatType.NotContainsText,
            "beginsWith" => ConditionalFormatType.BeginsWith,
            "endsWith" => ConditionalFormatType.EndsWith,
            "timePeriod" => ConditionalFormatType.ContainsDate,
            "top10" => ConditionalFormatType.Top10,
            "duplicateValues" => ConditionalFormatType.UniqueValues,
            "uniqueValues" => ConditionalFormatType.UniqueValues,
            "expression" => ConditionalFormatType.Expression,
            "colorScale" => ConditionalFormatType.ColorScale,
            "dataBar" => ConditionalFormatType.DataBar,
            "iconSet" => ConditionalFormatType.IconSet,
            "aboveAverage" => ConditionalFormatType.AboveAverage,
            "belowAverage" => ConditionalFormatType.BelowAverage,
            _ => ConditionalFormatType.CellIs
        };

        // 读取优先级
        var priority = reader.GetAttribute("priority");
        if (priority != null && int.TryParse(priority, out var pri))
        {
            format.Priority = pri;
        }

        // 读取停止标志
        var stopIfTrue = reader.GetAttribute("stopIfTrue");
        format.StopIfTrue = stopIfTrue == "1";

        // 读取操作符
        var op = reader.GetAttribute("operator");
        format.Operator = op switch
        {
            "between" => ComparisonOperator.Between,
            "notBetween" => ComparisonOperator.NotBetween,
            "equal" => ComparisonOperator.Equal,
            "notEqual" => ComparisonOperator.NotEqual,
            "greaterThan" => ComparisonOperator.GreaterThan,
            "lessThan" => ComparisonOperator.LessThan,
            "greaterThanOrEqual" => ComparisonOperator.GreaterThanOrEqual,
            "lessThanOrEqual" => ComparisonOperator.LessThanOrEqual,
            _ => ComparisonOperator.GreaterThan
        };

        if (reader.IsEmptyElement) return;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                switch (reader.LocalName)
                {
                    case "formula":
                        var formula = reader.ReadElementContentAsString();
                        if (string.IsNullOrEmpty(format.Formula1))
                        {
                            format.Formula1 = formula;
                        }
                        else
                        {
                            format.Formula2 = formula;
                        }
                        depth--;
                        break;
                    case "text":
                        format.Text = reader.GetAttribute("val");
                        break;
                    case "dxf":
                        format.Style = ReadDifferentialFormat(reader, ns);
                        depth--;
                        break;
                    case "colorScale":
                        format.ColorScale = ReadColorScale(reader, ns);
                        depth--;
                        break;
                    case "dataBar":
                        format.DataBar = ReadDataBar(reader, ns);
                        depth--;
                        break;
                    case "iconSet":
                        format.IconSet = ReadIconSet(reader, ns);
                        depth--;
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static List<CellRange> ParseRanges(string sqref)
    {
        var ranges = new List<CellRange>();
        var parts = sqref.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        foreach (var part in parts)
        {
            var range = ParseRange(part);
            if (range != null) ranges.Add(range);
        }

        return ranges;
    }

    private static CellRange? ParseRange(string range)
    {
        var colon = range.IndexOf(':');
        if (colon >= 0)
        {
            var left = range[..colon];
            var right = range[(colon + 1)..];
            ParseCellRef(left, out var firstRow, out var firstCol);
            ParseCellRef(right, out var lastRow, out var lastCol);
            return new CellRange { FirstRow = firstRow, FirstCol = firstCol, LastRow = lastRow, LastCol = lastCol };
        }
        else
        {
            ParseCellRef(range, out var row, out var col);
            return new CellRange { FirstRow = row, FirstCol = col, LastRow = row, LastCol = col };
        }
    }

    private static void ParseCellRef(string s, out int row, out int col)
    {
        row = col = 0;
        if (string.IsNullOrEmpty(s)) return;

        var i = 0;
        while (i < s.Length && char.IsLetter(s[i])) i++;
        if (i == 0 || i >= s.Length) return;

        if (!int.TryParse(s.AsSpan(i), out var r)) return;
        row = r - 1;
        col = ParseCol(s.AsSpan(0, i));
    }

    private static int ParseCol(ReadOnlySpan<char> s)
    {
        var col = 0;
        foreach (var c in s)
            col = col * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        return col - 1;
    }

    private static ConditionalFormatStyle? ReadDifferentialFormat(XmlReader reader, string ns)
    {
        var style = new ConditionalFormatStyle();
        var aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";

        if (reader.IsEmptyElement) return style;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.LocalName == "font" && reader.NamespaceURI == ns)
                {
                    ReadFontStyle(reader, ns, style);
                }
                else if (reader.LocalName == "fill" && reader.NamespaceURI == ns)
                {
                    ReadFillStyle(reader, ns, style);
                }
                else if (reader.LocalName == "border" && reader.NamespaceURI == ns)
                {
                    ReadBorderStyle(reader, ns, style);
                }
                else if (reader.LocalName == "numFmt" && reader.NamespaceURI == ns)
                {
                    style.NumberFormat = reader.GetAttribute("formatCode");
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }

        return style;
    }

    private static void ReadFontStyle(XmlReader reader, string ns, ConditionalFormatStyle style)
    {
        if (reader.IsEmptyElement) return;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                switch (reader.LocalName)
                {
                    case "b":
                        style.Bold = true;
                        break;
                    case "i":
                        style.Italic = true;
                        break;
                    case "color":
                        var color = ReadColor(reader, ns);
                        if (color.HasValue) style.FontColor = color.Value;
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static void ReadFillStyle(XmlReader reader, string ns, ConditionalFormatStyle style)
    {
        if (reader.IsEmptyElement) return;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "patternFill")
                {
                    var patternType = reader.GetAttribute("patternType");
                    if (patternType == "solid")
                    {
                        var innerDepth = 1;
                        while (reader.Read() && innerDepth > 0)
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
                            {
                                if (reader.LocalName == "fgColor")
                                {
                                    var color = ReadColor(reader, ns);
                                    if (color.HasValue) style.FillColor = color.Value;
                                }
                            }
                            else if (reader.NodeType == XmlNodeType.EndElement) innerDepth--;
                        }
                    }
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static void ReadBorderStyle(XmlReader reader, string ns, ConditionalFormatStyle style)
    {
        if (reader.IsEmptyElement) return;

        style.Border = new ConditionalFormatBorder();
        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                ChartColor? color = null;
                var innerDepth = 1;
                while (reader.Read() && innerDepth > 0)
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
                    {
                        if (reader.LocalName == "color")
                        {
                            color = ReadColor(reader, ns);
                        }
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement) innerDepth--;
                }

                switch (reader.LocalName)
                {
                    case "top":
                        if (color.HasValue) style.Border.TopColor = color.Value;
                        break;
                    case "bottom":
                        if (color.HasValue) style.Border.BottomColor = color.Value;
                        break;
                    case "left":
                        if (color.HasValue) style.Border.LeftColor = color.Value;
                        break;
                    case "right":
                        if (color.HasValue) style.Border.RightColor = color.Value;
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static ChartColor? ReadColor(XmlReader reader, string ns)
    {
        var rgb = reader.GetAttribute("rgb");
        if (!string.IsNullOrEmpty(rgb) && rgb.Length >= 6)
        {
            if (byte.TryParse(rgb[..2], System.Globalization.NumberStyles.HexNumber, null, out var r) &&
                byte.TryParse(rgb.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, null, out var g) &&
                byte.TryParse(rgb.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, null, out var b))
            {
                return new ChartColor(r, g, b);
            }
        }

        var theme = reader.GetAttribute("theme");
        if (theme != null)
        {
            // 主题颜色映射
            return theme switch
            {
                "0" => ChartColor.White,
                "1" => ChartColor.Black,
                "2" => ChartColor.Red,
                "3" => ChartColor.Green,
                "4" => ChartColor.Blue,
                "5" => ChartColor.Yellow,
                "6" => ChartColor.Cyan,
                "7" => ChartColor.Magenta,
                _ => ChartColor.Black
            };
        }

        return null;
    }

    private static ColorScale? ReadColorScale(XmlReader reader, string ns)
    {
        var colorScale = new ColorScale();
        var cfvoList = new List<(ColorScaleValueType type, double? value, string? formula)>();
        var colorList = new List<ChartColor>();

        if (reader.IsEmptyElement) return colorScale;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "cfvo")
                {
                    var type = reader.GetAttribute("type");
                    var val = reader.GetAttribute("val");

                    var valueType = type switch
                    {
                        "num" => ColorScaleValueType.Num,
                        "min" => ColorScaleValueType.MinValue,
                        "max" => ColorScaleValueType.MaxValue,
                        "percent" => ColorScaleValueType.Percent,
                        "percentile" => ColorScaleValueType.Percentile,
                        "formula" => ColorScaleValueType.Formula,
                        _ => ColorScaleValueType.MinValue
                    };

                    double? value = null;
                    if (val != null && double.TryParse(val, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var v))
                    {
                        value = v;
                    }

                    cfvoList.Add((valueType, value, null));
                }
                else if (reader.LocalName == "color")
                {
                    var color = ReadColor(reader, ns);
                    if (color.HasValue) colorList.Add(color.Value);
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }

        // 构建色阶
        if (cfvoList.Count >= 2 && colorList.Count >= 2)
        {
            colorScale.Minimum = new ColorScalePoint
            {
                Type = cfvoList[0].type,
                Value = cfvoList[0].value,
                Color = colorList[0]
            };

            colorScale.Maximum = new ColorScalePoint
            {
                Type = cfvoList[^1].type,
                Value = cfvoList[^1].value,
                Color = colorList[^1]
            };

            if (cfvoList.Count >= 3 && colorList.Count >= 3)
            {
                colorScale.Midpoint = new ColorScalePoint
                {
                    Type = cfvoList[1].type,
                    Value = cfvoList[1].value,
                    Color = colorList[1]
                };
            }
        }

        return colorScale;
    }

    private static DataBar? ReadDataBar(XmlReader reader, string ns)
    {
        var dataBar = new DataBar();

        // 读取显示值标志
        var showValue = reader.GetAttribute("showValue");
        dataBar.ShowValue = showValue != "0";

        if (reader.IsEmptyElement) return dataBar;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                switch (reader.LocalName)
                {
                    case "cfvo":
                        var type = reader.GetAttribute("type");
                        var val = reader.GetAttribute("val");

                        var valueType = type switch
                        {
                            "num" => DataBarValueType.Num,
                            "min" => DataBarValueType.MinValue,
                            "max" => DataBarValueType.MaxValue,
                            "percentile" => DataBarValueType.Percentile,
                            "formula" => DataBarValueType.Formula,
                            "auto" => DataBarValueType.Auto,
                            _ => DataBarValueType.MinValue
                        };

                        double? value = null;
                        if (val != null && double.TryParse(val, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var v))
                        {
                            value = v;
                        }

                        if (dataBar.Minimum.Type == DataBarValueType.MinValue && dataBar.Minimum.Value == null)
                        {
                            dataBar.Minimum = new DataBarPoint { Type = valueType, Value = value };
                        }
                        else
                        {
                            dataBar.Maximum = new DataBarPoint { Type = valueType, Value = value };
                        }
                        break;

                    case "color":
                        var color = ReadColor(reader, ns);
                        if (color.HasValue) dataBar.Color = color.Value;
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }

        return dataBar;
    }

    private static IconSet? ReadIconSet(XmlReader reader, string ns)
    {
        var iconSet = new IconSet();

        // 读取图标集类型
        var setType = reader.GetAttribute("iconSet");
        iconSet.Type = setType switch
        {
            "3Arrows" => IconSetType.ThreeArrows,
            "3ArrowsGray" => IconSetType.ThreeArrowsGray,
            "3Flags" => IconSetType.ThreeFlags,
            "3TrafficLights1" => IconSetType.ThreeTrafficLights,
            "3TrafficLights2" => IconSetType.ThreeSigns,
            "3Signs" => IconSetType.ThreeSigns,
            "3Symbols" => IconSetType.ThreeSymbols,
            "3Symbols2" => IconSetType.ThreeSymbols2,
            "4Arrows" => IconSetType.FourArrows,
            "4ArrowsGray" => IconSetType.FourArrowsGray,
            "4RedToBlack" => IconSetType.FourRedToBlack,
            "4Rating" => IconSetType.FourRatings,
            "4TrafficLights" => IconSetType.FourTrafficLights,
            "5Arrows" => IconSetType.FiveArrows,
            "5ArrowsGray" => IconSetType.FiveArrowsGray,
            "5Rating" => IconSetType.FiveRatings,
            "5Quarters" => IconSetType.FiveQuarters,
            "3Stars" => IconSetType.ThreeStars,
            "3Triangles" => IconSetType.ThreeTriangles,
            "5Boxes" => IconSetType.FiveBoxes,
            _ => IconSetType.ThreeTrafficLights
        };

        // 读取显示值标志
        var showValue = reader.GetAttribute("showValue");
        iconSet.ShowValue = showValue != "0";

        // 读取反转标志
        var reverse = reader.GetAttribute("reverse");
        iconSet.Reverse = reverse == "1";

        if (reader.IsEmptyElement) return iconSet;

        var thresholds = new List<IconThreshold>();
        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "cfvo")
                {
                    var type = reader.GetAttribute("type");
                    var val = reader.GetAttribute("val");

                    var valueType = type switch
                    {
                        "num" => IconValueType.Num,
                        "percent" => IconValueType.Percent,
                        "formula" => IconValueType.Formula,
                        "percentile" => IconValueType.Percentile,
                        _ => IconValueType.Percent
                    };

                    double value = 0;
                    if (val != null && double.TryParse(val, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var v))
                    {
                        value = v;
                    }

                    thresholds.Add(new IconThreshold
                    {
                        Type = valueType,
                        Value = value
                    });
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }

        if (thresholds.Count > 0)
        {
            iconSet.Thresholds = thresholds;
        }

        return iconSet;
    }
}
