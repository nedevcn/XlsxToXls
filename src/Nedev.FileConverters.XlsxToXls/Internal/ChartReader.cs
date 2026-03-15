using System.IO.Compression;
using System.Xml;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// XLSX图表读取器 - 从Open XML格式读取图表数据
/// </summary>
internal static class ChartReader
{
    private static readonly XmlReaderSettings Settings = new()
    {
        Async = false,
        CloseInput = false,
        IgnoreWhitespace = true,
        DtdProcessing = DtdProcessing.Prohibit
    };

    public static List<ChartData> ReadCharts(ZipArchive archive, string worksheetPath, Action<string>? log = null)
    {
        var charts = new List<ChartData>();
        try
        {
            var drawingPath = FindDrawingPath(archive, worksheetPath);
            if (string.IsNullOrEmpty(drawingPath))
            {
                log?.Invoke($"[ChartReader] No drawing found for worksheet: {worksheetPath}");
                return charts;
            }

            log?.Invoke($"[ChartReader] Found drawing: {drawingPath}");

            var chartRefs = ReadDrawingForCharts(archive, drawingPath);
            log?.Invoke($"[ChartReader] Found {chartRefs.Count} chart references");

            foreach (var chartPath in chartRefs)
            {
                try
                {
                    var chart = ReadChartFile(archive, chartPath, log);
                    if (chart != null)
                    {
                        charts.Add(chart);
                        log?.Invoke($"[ChartReader] Successfully read chart: {chart.Name} ({chart.Type})");
                    }
                }
                catch (Exception ex)
                {
                    log?.Invoke($"[ChartReader] Error reading chart from {chartPath}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[ChartReader] Error reading charts: {ex.Message}");
        }
        return charts;
    }

    private static string? FindDrawingPath(ZipArchive archive, string worksheetPath)
    {
        var relsPath = worksheetPath.Replace("worksheets/", "worksheets/_rels/") + ".rels";
        var entry = archive.GetEntry(relsPath) ?? archive.GetEntry("xl/" + relsPath);
        if (entry == null) return null;

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Settings);
        var ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship" && reader.NamespaceURI == ns)
            {
                var type = reader.GetAttribute("Type");
                var target = reader.GetAttribute("Target");
                if (type?.EndsWith("/drawing", StringComparison.OrdinalIgnoreCase) == true && target != null)
                {
                    return target.StartsWith("/") ? target.TrimStart('/') : "xl/drawings/" + target.Split('/').Last();
                }
            }
        }
        return null;
    }

    private static List<string> ReadDrawingForCharts(ZipArchive archive, string drawingPath)
    {
        var chartPaths = new List<string>();
        var entry = archive.GetEntry(drawingPath) ?? archive.GetEntry("xl/drawings/" + drawingPath.Split('/').Last());
        if (entry == null) return chartPaths;

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Settings);
        var chartNs = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "chart" && reader.NamespaceURI == chartNs)
                {
                var id = reader.GetAttribute("id");
                if (!string.IsNullOrEmpty(id))
                {
                    var chartPath = ResolveChartPath(archive, drawingPath, id);
                    if (!string.IsNullOrEmpty(chartPath)) chartPaths.Add(chartPath);
                }
            }
        }
        return chartPaths;
    }

    private static string? ResolveChartPath(ZipArchive archive, string drawingPath, string rId)
    {
        var relsPath = drawingPath.Replace("drawings/", "drawings/_rels/") + ".rels";
        var entry = archive.GetEntry(relsPath) ?? archive.GetEntry("xl/" + relsPath);
        if (entry == null) return null;

        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Settings);
        var ns = "http://schemas.openxmlformats.org/package/2006/relationships";

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship" && reader.NamespaceURI == ns)
            {
                var id = reader.GetAttribute("Id");
                var type = reader.GetAttribute("Type");
                var target = reader.GetAttribute("Target");
                if (id == rId && type?.EndsWith("/chart", StringComparison.OrdinalIgnoreCase) == true && target != null)
                {
                    return target.StartsWith("/") ? target.TrimStart('/') : "xl/charts/" + target.Split('/').Last();
                }
            }
        }
        return null;
    }

    private static ChartData? ReadChartFile(ZipArchive archive, string chartPath, Action<string>? log = null)
    {
        var entry = archive.GetEntry(chartPath) ?? archive.GetEntry("xl/charts/" + chartPath.Split('/').Last());
        if (entry == null)
        {
            log?.Invoke($"[ChartReader] Chart file not found: {chartPath}");
            return null;
        }

        log?.Invoke($"[ChartReader] Reading chart file: {chartPath}");

        var chart = new ChartData();
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream, Settings);
        var ns = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element) continue;

            switch (reader.LocalName)
            {
                case "barChart":
                    chart.Type = ChartType.Bar;
                    ReadBarChart(reader, ns, chart);
                    break;
                case "lineChart":
                    chart.Type = ChartType.Line;
                    ReadLineChart(reader, ns, chart);
                    break;
                case "pieChart":
                case "pie3DChart":
                    chart.Type = ChartType.Pie;
                    ReadPieChart(reader, ns, chart);
                    break;
                case "areaChart":
                    chart.Type = ChartType.Area;
                    ReadAreaChart(reader, ns, chart);
                    break;
                case "scatterChart":
                    chart.Type = ChartType.Scatter;
                    ReadScatterChart(reader, ns, chart);
                    break;
                case "radarChart":
                    chart.Type = ChartType.Radar;
                    ReadRadarChart(reader, ns, chart);
                    break;
                case "doughnutChart":
                    chart.Type = ChartType.Doughnut;
                    ReadDoughnutChart(reader, ns, chart);
                    break;
                case "bubbleChart":
                    chart.Type = ChartType.Bubble;
                    ReadBubbleChart(reader, ns, chart);
                    break;
                case "surfaceChart":
                case "surface3DChart":
                    chart.Type = ChartType.Surface;
                    ReadSurfaceChart(reader, ns, chart);
                    break;
                case "stockChart":
                    chart.Type = ChartType.StockOHLC;
                    ReadStockChart(reader, ns, chart);
                    break;
                case "area3DChart":
                    chart.Type = ChartType.Area;
                    ReadAreaChart(reader, ns, chart);
                    break;
                case "bar3DChart":
                    chart.Type = ChartType.Bar;
                    ReadBarChart(reader, ns, chart);
                    break;
                case "line3DChart":
                    chart.Type = ChartType.Line;
                    ReadLineChart(reader, ns, chart);
                    break;
                case "column3DChart":
                    chart.Type = ChartType.Column;
                    ReadColumnChart(reader, ns, chart);
                    break;
                case "cylinderChart":
                    chart.Type = ChartType.CylinderColumn;
                    ReadColumnChart(reader, ns, chart);
                    break;
                case "coneChart":
                    chart.Type = ChartType.ConeColumn;
                    ReadColumnChart(reader, ns, chart);
                    break;
                case "pyramidChart":
                    chart.Type = ChartType.PyramidColumn;
                    ReadColumnChart(reader, ns, chart);
                    break;
                case "title":
                    chart.Title = ReadTitle(reader, ns);
                    break;
                case "legend":
                    chart.Legend = ReadLegend(reader, ns);
                    break;
                case "catAx":
                    chart.CategoryAxis = ReadAxis(reader, ns, AxisType.Category);
                    break;
                case "valAx":
                    // 检查是否是次坐标轴
                    var axisId = reader.GetAttribute("axId");
                    if (chart.ValueAxis == null)
                    {
                        chart.ValueAxis = ReadAxis(reader, ns, AxisType.Value);
                    }
                    else
                    {
                        chart.SecondaryValueAxis = ReadAxis(reader, ns, AxisType.Value);
                    }
                    break;
                case "dTable":
                    chart.DataTable = ReadDataTable(reader, ns);
                    break;
                case "plotArea":
                    // 检查是否是组合图表
                    CheckForComboChart(reader, ns, chart);
                    break;
            }
        }

        return chart.Series.Count > 0 ? chart : null;
    }

    private static void ReadBarChart(XmlReader reader, string ns, ChartData chart)
    {
        if (reader.IsEmptyElement) return;
        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                depth++;
                if (reader.LocalName == "ser" && reader.NamespaceURI == ns)
                {
                    var series = ReadSeries(reader, ns);
                    if (series != null) chart.Series.Add(series);
                    depth--;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static void ReadLineChart(XmlReader reader, string ns, ChartData chart)
    {
        ReadBarChart(reader, ns, chart);
    }

    private static void ReadPieChart(XmlReader reader, string ns, ChartData chart)
    {
        ReadBarChart(reader, ns, chart);
    }

    private static void ReadAreaChart(XmlReader reader, string ns, ChartData chart)
    {
        ReadBarChart(reader, ns, chart);
    }

    private static void ReadScatterChart(XmlReader reader, string ns, ChartData chart)
    {
        ReadBarChart(reader, ns, chart);
    }

    private static void ReadRadarChart(XmlReader reader, string ns, ChartData chart)
    {
        if (reader.IsEmptyElement) return;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                switch (reader.LocalName)
                {
                    case "radarStyle":
                        var style = reader.GetAttribute("val");
                        chart.PlotArea.RadarStyle = style?.ToLowerInvariant() switch
                        {
                            "filled" => RadarStyle.Filled,
                            _ => RadarStyle.Marker
                        };
                        break;
                    case "axId":
                        // Radar chart axis handling
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static void ReadDoughnutChart(XmlReader reader, string ns, ChartData chart)
    {
        if (reader.IsEmptyElement) return;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                switch (reader.LocalName)
                {
                    case "firstSliceAng":
                        if (int.TryParse(reader.GetAttribute("val"), out var angle))
                            chart.PlotArea.FirstSliceAngle = angle;
                        break;
                    case "holeSize":
                        if (int.TryParse(reader.GetAttribute("val"), out var holeSize))
                            chart.PlotArea.HoleSize = holeSize;
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static void ReadBubbleChart(XmlReader reader, string ns, ChartData chart)
    {
        if (reader.IsEmptyElement) return;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                switch (reader.LocalName)
                {
                    case "bubbleScale":
                        if (int.TryParse(reader.GetAttribute("val"), out var scale))
                            chart.PlotArea.BubbleScale = scale;
                        break;
                    case "showNegBubbles":
                        chart.PlotArea.ShowNegativeBubbles = reader.GetAttribute("val") == "1";
                        break;
                    case "sizeRepresents":
                        // Size represents area or width
                        break;
                    case "axId":
                        // Bubble chart axis handling
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static void ReadSurfaceChart(XmlReader reader, string ns, ChartData chart)
    {
        if (reader.IsEmptyElement) return;

        chart.PlotArea.SurfaceViewSettings ??= new SurfaceViewSettings();

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                switch (reader.LocalName)
                {
                    case "wireframe":
                        if (reader.GetAttribute("val") == "1")
                            chart.Type = ChartType.SurfaceWireframe;
                        break;
                    case "axId":
                        // Surface chart axis handling
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static void ReadStockChart(XmlReader reader, string ns, ChartData chart)
    {
        if (reader.IsEmptyElement) return;

        chart.PlotArea.StockSettings ??= new StockSettings();

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                switch (reader.LocalName)
                {
                    case "dropLines":
                        chart.PlotArea.StockSettings.ShowDropLines = true;
                        break;
                    case "highLowLines":
                        chart.PlotArea.StockSettings.ShowHighLowLines = true;
                        break;
                    case "upDownBars":
                        chart.PlotArea.StockSettings.ShowOpenCloseBars = true;
                        break;
                    case "axId":
                        // Stock chart axis handling
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static void ReadColumnChart(XmlReader reader, string ns, ChartData chart)
    {
        ReadBarChart(reader, ns, chart);
    }

    private static ChartSeries? ReadSeries(XmlReader reader, string ns)
    {
        var series = new ChartSeries();
        if (reader.IsEmptyElement) return series;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                depth++;
                switch (reader.LocalName)
                {
                    case "tx":
                        series.Name = ReadSeriesText(reader, ns);
                        depth--;
                        break;
                    case "cat":
                        case "xVal":
                        series.Categories = ReadChartReference(reader, ns);
                        depth--;
                        break;
                    case "val":
                    case "yVal":
                        series.Values = ReadChartReference(reader, ns);
                        depth--;
                        break;
                    case "dLbls":
                        series.DataLabels = ReadDataLabels(reader, ns);
                        depth--;
                        break;
                    case "spPr":
                        ReadSeriesStyle(reader, ns, series);
                        depth--;
                        break;
                    case "dPt":
                        var point = ReadDataPoint(reader, ns);
                        if (point != null)
                        {
                            series.DataPoints ??= new List<ChartDataPoint>();
                            series.DataPoints.Add(point);
                        }
                        depth--;
                        break;
                    case "trendline":
                        var trendLine = ReadTrendLine(reader, ns);
                        if (trendLine != null)
                        {
                            series.TrendLines ??= new List<TrendLine>();
                            series.TrendLines.Add(trendLine);
                        }
                        depth--;
                        break;
                    case "errBars":
                        series.ErrorBars = ReadErrorBars(reader, ns);
                        depth--;
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return series;
    }

    private static ChartDataPoint? ReadDataPoint(XmlReader reader, string ns)
    {
        var point = new ChartDataPoint();
        var idx = reader.GetAttribute("idx");
        if (idx != null && int.TryParse(idx, out var index))
        {
            point.Index = index;
        }

        if (reader.IsEmptyElement) return point;

        var aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.LocalName == "solidFill" && reader.NamespaceURI == aNs)
                {
                    var color = ReadColor(reader, aNs);
                    if (color.HasValue) point.FillColor = color.Value;
                }
                else if (reader.LocalName == "ln" && reader.NamespaceURI == aNs)
                {
                    var color = ReadColor(reader, aNs);
                    if (color.HasValue) point.BorderColor = color.Value;
                }
                else if (reader.LocalName == "dLbl" && reader.NamespaceURI == ns)
                {
                    point.DataLabels = ReadDataLabels(reader, ns);
                }
                else if (reader.LocalName == "explosion" && reader.NamespaceURI == ns)
                {
                    var val = reader.GetAttribute("val");
                    point.Explosion = val != null && int.TryParse(val, out var exp) && exp > 0;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return point;
    }

    private static TrendLine? ReadTrendLine(XmlReader reader, string ns)
    {
        var trendLine = new TrendLine();

        // 读取趋势线类型
        var type = reader.GetAttribute("trendlineType");
        trendLine.Type = type switch
        {
            "linear" => TrendLineType.Linear,
            "exp" => TrendLineType.Exponential,
            "log" => TrendLineType.Logarithmic,
            "poly" => TrendLineType.Polynomial,
            "power" => TrendLineType.Power,
            "movingAvg" => TrendLineType.MovingAverage,
            _ => TrendLineType.Linear
        };

        // 读取阶数（多项式）
        var order = reader.GetAttribute("order");
        if (order != null && int.TryParse(order, out var ord))
        {
            trendLine.Order = ord;
        }

        // 读取周期（移动平均）
        var period = reader.GetAttribute("period");
        if (period != null && int.TryParse(period, out var per))
        {
            trendLine.Period = per;
        }

        if (reader.IsEmptyElement) return trendLine;

        var aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.LocalName == "trendlineLbl" && reader.NamespaceURI == ns)
                {
                    var depth2 = 1;
                    while (reader.Read() && depth2 > 0)
                    {
                        if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
                        {
                            if (reader.LocalName == "showEq")
                            {
                                var val = reader.GetAttribute("val");
                                trendLine.DisplayEquation = val == "1";
                            }
                            else if (reader.LocalName == "showRSqrVal")
                            {
                                var val = reader.GetAttribute("val");
                                trendLine.DisplayRSquared = val == "1";
                            }
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement) depth2--;
                    }
                }
                else if (reader.LocalName == "spPr" && reader.NamespaceURI == aNs)
                {
                    var depth2 = 1;
                    while (reader.Read() && depth2 > 0)
                    {
                        if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == aNs)
                        {
                            if (reader.LocalName == "solidFill")
                            {
                                var color = ReadColor(reader, aNs);
                                if (color.HasValue) trendLine.LineColor = color.Value;
                            }
                            else if (reader.LocalName == "prstDash")
                            {
                                var val = reader.GetAttribute("val");
                                trendLine.LineStyle = val switch
                                {
                                    "solid" => LineStyle.Solid,
                                    "dash" => LineStyle.Dash,
                                    "dot" => LineStyle.Dot,
                                    "dashDot" => LineStyle.DashDot,
                                    _ => LineStyle.Solid
                                };
                            }
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement) depth2--;
                    }
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return trendLine;
    }

    private static ErrorBars? ReadErrorBars(XmlReader reader, string ns)
    {
        var errorBars = new ErrorBars();

        // 读取误差线方向
        var errDir = reader.GetAttribute("errDir");
        if (errDir == "X")
        {
            // X轴误差线
        }

        // 读取误差线类型
        var errBarType = reader.GetAttribute("errBarType");
        errorBars.Type = errBarType switch
        {
            "both" => ErrorBarType.Both,
            "plus" => ErrorBarType.Plus,
            "minus" => ErrorBarType.Minus,
            _ => ErrorBarType.Both
        };

        if (reader.IsEmptyElement) return errorBars;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "errValType")
                {
                    var val = reader.GetAttribute("val");
                    errorBars.ValueType = val switch
                    {
                        "fixedVal" => ErrorBarValueType.FixedValue,
                        "percentage" => ErrorBarValueType.Percentage,
                        "stdDev" => ErrorBarValueType.StandardDeviation,
                        "stdErr" => ErrorBarValueType.StandardError,
                        "cust" => ErrorBarValueType.Custom,
                        _ => ErrorBarValueType.FixedValue
                    };
                }
                else if (reader.LocalName == "val")
                {
                    var val = reader.GetAttribute("val");
                    if (val != null && double.TryParse(val, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var v))
                    {
                        errorBars.Value = v;
                    }
                }
                else if (reader.LocalName == "noEndCap")
                {
                    var val = reader.GetAttribute("val");
                    errorBars.ShowCap = val != "1";
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return errorBars;
    }

    private static DataLabels? ReadDataLabels(XmlReader reader, string ns)
    {
        var labels = new DataLabels();
        if (reader.IsEmptyElement) return labels;

        // 检查是否有 showVal 属性
        var showVal = reader.GetAttribute("showVal");
        if (showVal != null) labels.ShowValue = showVal == "1";

        // 检查是否有 showCatName 属性
        var showCatName = reader.GetAttribute("showCatName");
        if (showCatName != null) labels.ShowCategory = showCatName == "1";

        // 检查是否有 showPercent 属性
        var showPercent = reader.GetAttribute("showPercent");
        if (showPercent != null) labels.ShowPercentage = showPercent == "1";

        // 检查是否有 showSerName 属性
        var showSerName = reader.GetAttribute("showSerName");
        if (showSerName != null) labels.ShowSeriesName = showSerName == "1";

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                switch (reader.LocalName)
                {
                    case "dLblPos":
                        var val = reader.GetAttribute("val");
                        labels.Position = val switch
                        {
                            "ctr" => DataLabelPosition.Center,
                            "inEnd" => DataLabelPosition.InsideEnd,
                            "outEnd" => DataLabelPosition.OutsideEnd,
                            "bestFit" => DataLabelPosition.BestFit,
                            "l" => DataLabelPosition.Left,
                            "r" => DataLabelPosition.Right,
                            "t" => DataLabelPosition.Above,
                            "b" => DataLabelPosition.Below,
                            _ => DataLabelPosition.OutsideEnd
                        };
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return labels;
    }

    private static void ReadSeriesStyle(XmlReader reader, string ns, ChartSeries series)
    {
        if (reader.IsEmptyElement) return;

        var aNs = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.LocalName == "solidFill" && reader.NamespaceURI == aNs)
                {
                    // 读取填充颜色
                    var color = ReadColor(reader, aNs);
                    if (color.HasValue) series.FillColor = color.Value;
                }
                else if (reader.LocalName == "ln" && reader.NamespaceURI == aNs)
                {
                    // 读取线条样式
                    ReadLineStyle(reader, aNs, series);
                }
                else if (reader.LocalName == "marker" && reader.NamespaceURI == ns)
                {
                    // 读取标记样式
                    ReadMarkerStyle(reader, ns, series);
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static ChartColor? ReadColor(XmlReader reader, string ns)
    {
        if (reader.IsEmptyElement) return null;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "srgbClr")
                {
                    var val = reader.GetAttribute("val");
                    if (!string.IsNullOrEmpty(val) && val.Length >= 6)
                    {
                        if (byte.TryParse(val[..2], System.Globalization.NumberStyles.HexNumber, null, out var r) &&
                            byte.TryParse(val.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, null, out var g) &&
                            byte.TryParse(val.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, null, out var b))
                        {
                            return new ChartColor(r, g, b);
                        }
                    }
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return null;
    }

    private static void ReadLineStyle(XmlReader reader, string ns, ChartSeries series)
    {
        if (reader.IsEmptyElement) return;

        var width = reader.GetAttribute("w");
        if (width != null && int.TryParse(width, out var w))
        {
            // width 是以 EMU 为单位，转换为线条样式
            series.LineStyle = w > 0 ? LineStyle.Solid : LineStyle.None;
        }

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "solidFill")
                {
                    var color = ReadColor(reader, ns);
                    if (color.HasValue) series.BorderColor = color.Value;
                }
                else if (reader.LocalName == "noFill")
                {
                    series.LineStyle = LineStyle.None;
                }
                else if (reader.LocalName == "prstDash")
                {
                    var val = reader.GetAttribute("val");
                    series.LineStyle = val switch
                    {
                        "solid" => LineStyle.Solid,
                        "dash" => LineStyle.Dash,
                        "dot" => LineStyle.Dot,
                        "dashDot" => LineStyle.DashDot,
                        "lgDash" => LineStyle.Dash,
                        _ => LineStyle.Solid
                    };
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static void ReadMarkerStyle(XmlReader reader, string ns, ChartSeries series)
    {
        if (reader.IsEmptyElement) return;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "symbol")
                {
                    var val = reader.GetAttribute("val");
                    series.MarkerStyle = val switch
                    {
                        "square" => MarkerStyle.Square,
                        "diamond" => MarkerStyle.Diamond,
                        "triangle" => MarkerStyle.Triangle,
                        "x" => MarkerStyle.X,
                        "star" => MarkerStyle.Star,
                        "dot" => MarkerStyle.Dot,
                        "circle" => MarkerStyle.Circle,
                        "plus" => MarkerStyle.Plus,
                        "none" => MarkerStyle.None,
                        _ => MarkerStyle.None
                    };
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static string? ReadSeriesText(XmlReader reader, string ns)
    {
        if (reader.IsEmptyElement) return null;
        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "v")
                {
                    if (reader.Read() && reader.NodeType == XmlNodeType.Text)
                        return reader.Value;
                }
                else if (reader.LocalName == "f")
                {
                    var formula = reader.ReadElementContentAsString();
                    return formula;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return null;
    }

    private static ChartRange? ReadChartReference(XmlReader reader, string ns)
    {
        if (reader.IsEmptyElement) return new ChartRange();

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "f")
                {
                    var formula = reader.ReadElementContentAsString();
                    return ParseFormula(formula);
                }
                else if (reader.LocalName == "strRef" || reader.LocalName == "numRef")
                {
                    var innerDepth = 1;
                    while (reader.Read() && innerDepth > 0)
                    {
                        if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns && reader.LocalName == "f")
                        {
                            var formula = reader.ReadElementContentAsString();
                            return ParseFormula(formula);
                        }
                        else if (reader.NodeType == XmlNodeType.EndElement) innerDepth--;
                    }
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return new ChartRange();
    }

    private static ChartRange ParseFormula(string formula)
    {
        var sheetName = "";
        var excl = formula.IndexOf('!');
        if (excl >= 0)
        {
            sheetName = formula[..excl].Trim('\'');
            formula = formula[(excl + 1)..];
        }

        formula = formula.Replace("$", "");
        var colon = formula.IndexOf(':');
        if (colon >= 0)
        {
            var left = formula[..colon];
            var right = formula[(colon + 1)..];
            ParseCellRef(left, out var firstRow, out var firstCol);
            ParseCellRef(right, out var lastRow, out var lastCol);
            return new ChartRange { SheetName = sheetName, FirstRow = firstRow, FirstCol = firstCol, LastRow = lastRow, LastCol = lastCol };
        }
        else
        {
            ParseCellRef(formula, out var row, out var col);
            return new ChartRange { SheetName = sheetName, FirstRow = row, FirstCol = col, LastRow = row, LastCol = col };
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

    private static ChartTitle? ReadTitle(XmlReader reader, string ns)
    {
        var title = new ChartTitle();
        if (reader.IsEmptyElement) return title;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "v")
                {
                    if (reader.Read() && reader.NodeType == XmlNodeType.Text)
                        title.Text = reader.Value;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return title;
    }

    private static ChartLegend? ReadLegend(XmlReader reader, string ns)
    {
        var legend = new ChartLegend();
        var pos = reader.GetAttribute("legendPos");
        if (!string.IsNullOrEmpty(pos))
        {
            legend.Position = pos.ToLowerInvariant() switch
            {
                "l" => LegendPosition.Left,
                "r" => LegendPosition.Right,
                "t" => LegendPosition.Top,
                "b" => LegendPosition.Bottom,
                "tr" => LegendPosition.Corner,
                _ => LegendPosition.Right
            };
        }
        return legend;
    }

    private static ChartAxis ReadAxis(XmlReader reader, string ns, AxisType type)
    {
        var axis = new ChartAxis { Type = type };
        if (reader.IsEmptyElement) return axis;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                switch (reader.LocalName)
                {
                    case "title":
                        var title = ReadTitle(reader, ns);
                        if (title != null) axis.Title = title.Text;
                        break;
                    case "scaling":
                        ReadScaling(reader, ns, axis);
                        break;
                    case "majorGridlines":
                        axis.HasMajorGridlines = true;
                        break;
                    case "minorGridlines":
                        axis.HasMinorGridlines = true;
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return axis;
    }

    private static void ReadScaling(XmlReader reader, string ns, ChartAxis axis)
    {
        if (reader.IsEmptyElement) return;
        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                if (reader.LocalName == "min")
                {
                    var val = reader.GetAttribute("val");
                    if (double.TryParse(val, out var min)) axis.MinValue = min;
                }
                else if (reader.LocalName == "max")
                {
                    var val = reader.GetAttribute("val");
                    if (double.TryParse(val, out var max)) axis.MaxValue = max;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }

    private static ChartDataTable? ReadDataTable(XmlReader reader, string ns)
    {
        var dataTable = new ChartDataTable();

        // 读取showHorzBorder属性
        var showHorzBorder = reader.GetAttribute("showHorzBorder");
        if (showHorzBorder != null) dataTable.HasHorizontalBorder = showHorzBorder == "1";

        // 读取showVertBorder属性
        var showVertBorder = reader.GetAttribute("showVertBorder");
        if (showVertBorder != null) dataTable.HasVerticalBorder = showVertBorder == "1";

        // 读取showOutline属性
        var showOutline = reader.GetAttribute("showOutline");
        if (showOutline != null) dataTable.HasOutlineBorder = showOutline == "1";

        // 读取showKeys属性
        var showKeys = reader.GetAttribute("showKeys");
        if (showKeys != null) dataTable.ShowLegendKeys = showKeys == "1";

        if (reader.IsEmptyElement) return dataTable;

        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
        return dataTable;
    }

    private static void CheckForComboChart(XmlReader reader, string ns, ChartData chart)
    {
        if (reader.IsEmptyElement) return;

        var chartTypeCount = 0;
        var depth = 1;
        while (reader.Read() && depth > 0)
        {
            if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == ns)
            {
                var localName = reader.LocalName;
                if (localName.EndsWith("Chart") && localName != "chart")
                {
                    chartTypeCount++;
                    if (chartTypeCount > 1)
                    {
                        chart.IsComboChart = true;
                    }
                }
                else if (localName == "axId")
                {
                    // 读取坐标轴ID以识别主次坐标轴
                    var val = reader.GetAttribute("val");
                    // 这里可以添加更复杂的逻辑来关联系列和坐标轴
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement) depth--;
        }
    }
}
