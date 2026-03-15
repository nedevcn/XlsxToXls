using System.Buffers;
using System.Buffers.Binary;
using System.Text;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// BIFF8图表记录写入器 - 使用ArrayPool减少内存分配
/// </summary>
internal ref struct ChartWriter
{
    private Span<byte> _buffer;
    private int _position;
    private byte[]? _pooledBuffer;

    public ChartWriter(Span<byte> buffer)
    {
        _buffer = buffer;
        _position = 0;
        _pooledBuffer = null;
    }

    /// <summary>
    /// 使用ArrayPool创建ChartWriter，自动管理缓冲区
    /// </summary>
    public static ChartWriter CreatePooled(out byte[] pooledBuffer, int minSize = 65536)
    {
        pooledBuffer = ArrayPool<byte>.Shared.Rent(minSize);
        return new ChartWriter(pooledBuffer.AsSpan())
        {
            _pooledBuffer = pooledBuffer
        };
    }

    /// <summary>
    /// 释放ArrayPool缓冲区（如果使用了CreatePooled）
    /// </summary>
    public void Dispose()
    {
        if (_pooledBuffer != null)
        {
            ArrayPool<byte>.Shared.Return(_pooledBuffer);
            _pooledBuffer = null;
        }
    }

    public int Position => _position;

    /// <summary>
    /// 写入完整的图表流
    /// </summary>
    public int WriteChartStream(ChartData chart, int sheetIndex)
    {
        WriteBofChart();
        WriteChartTypeRecord(chart.Type);

        // 写入图表标题
        if (!string.IsNullOrEmpty(chart.Title?.Text))
        {
            WriteChartTitle(chart.Title.Text);
        }

        // 写入图例
        if (chart.Legend?.Show == true)
        {
            WriteLegend(chart.Legend.Position);
        }

        // 写入绘图区
        WritePlotArea(chart);

        // 写入数据系列
        for (var i = 0; i < chart.Series.Count; i++)
        {
            WriteSeries(chart.Series[i], i);
        }

        // 写入坐标轴
        if (chart.CategoryAxis != null)
        {
            WriteAxis(chart.CategoryAxis);
        }
        if (chart.ValueAxis != null)
        {
            WriteAxis(chart.ValueAxis);
        }
        if (chart.SecondaryValueAxis != null)
        {
            WriteAxis(chart.SecondaryValueAxis);
        }

        // 写入系列到轴的关联
        WriteAxisLink();

        // 写入数据表
        if (chart.DataTable?.Show == true)
        {
            WriteDataTable(chart.DataTable);
        }

        WriteEof();
        return _position;
    }

    private void WriteBofChart()
    {
        WriteRecordHeader(0x0809, 16);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0600);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0020);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0C0A);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x07CC);
        _position += 2;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0x00000001);
        _position += 4;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0006);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0000);
        _position += 2;
    }

    private void WriteChartTypeRecord(ChartType type)
    {
        // CHARTTYPE记录 (0x1000系列)
        var recordType = type switch
        {
            ChartType.Area => 0x101A,
            ChartType.Bar => 0x1017,
            ChartType.Line => 0x1018,
            ChartType.Pie => 0x1019,
            ChartType.Scatter => 0x101B,
            ChartType.Radar => 0x103C,
            ChartType.RadarWithMarkers => 0x103C,
            ChartType.Column => 0x1017, // 柱状图使用Bar记录，通过标志位区分
            ChartType.Doughnut => 0x102C,
            ChartType.Bubble => 0x103E,
            ChartType.Surface => 0x103F,
            ChartType.SurfaceWireframe => 0x103F,
            ChartType.StockHLC => 0x1042,
            ChartType.StockOHLC => 0x1042,
            ChartType.StockVHLC => 0x1042,
            ChartType.StockVOHLC => 0x1042,
            ChartType.ConeColumn => 0x1017,
            ChartType.ConeBar => 0x1017,
            ChartType.CylinderColumn => 0x1017,
            ChartType.CylinderBar => 0x1017,
            ChartType.PyramidColumn => 0x1017,
            ChartType.PyramidBar => 0x1017,
            _ => 0x1017
        };

        WriteRecordHeader((ushort)recordType, 6);

        // 图表类型标志
        var flags = type switch
        {
            ChartType.Column => 0x0001,
            ChartType.ConeColumn => 0x0001,
            ChartType.CylinderColumn => 0x0001,
            ChartType.PyramidColumn => 0x0001,
            ChartType.ConeBar => 0x0002,
            ChartType.CylinderBar => 0x0002,
            ChartType.PyramidBar => 0x0002,
            ChartType.SurfaceWireframe => 0x0001,
            _ => 0x0000
        };
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)flags);
        _position += 2;

        // 预留字段
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;

        // 写入高级图表类型特定设置
        WriteAdvancedChartSettings(type);
    }

    private void WriteAdvancedChartSettings(ChartType type)
    {
        switch (type)
        {
            case ChartType.Bubble:
                WriteBubbleSettings();
                break;
            case ChartType.Radar:
            case ChartType.RadarWithMarkers:
                WriteRadarSettings();
                break;
            case ChartType.StockHLC:
            case ChartType.StockOHLC:
            case ChartType.StockVHLC:
            case ChartType.StockVOHLC:
                WriteStockSettings();
                break;
            case ChartType.Surface:
            case ChartType.SurfaceWireframe:
                WriteSurfaceSettings();
                break;
        }
    }

    private void WriteBubbleSettings()
    {
        // BUBBLESIZE记录 (0x1043)
        WriteRecordHeader(0x1043, 4);

        // 气泡缩放比例 (默认100%)
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 100);
        _position += 2;

        // 标志位
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteRadarSettings()
    {
        // RADARFLAGS记录 (0x1044)
        WriteRecordHeader(0x1044, 2);

        // 雷达图标志
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteStockSettings()
    {
        // STOCKFLAGS记录 (0x1045)
        WriteRecordHeader(0x1045, 4);

        // 股价图标志
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0x0001);
        _position += 4;
    }

    private void WriteSurfaceSettings()
    {
        // SURFACEFLAGS记录 (0x1046)
        WriteRecordHeader(0x1046, 2);

        // 曲面图标志
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteChartTitle(string title)
    {
        // CHARTTITLE记录 (0x102D)
        var bytes = Encoding.Unicode.GetBytes(title);
        var recLen = 4 + bytes.Length;

        WriteRecordHeader(0x102D, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)title.Length);
        _position += 2;
        _buffer[_position++] = 1; // Unicode标志
        bytes.CopyTo(_buffer.Slice(_position));
        _position += bytes.Length;
    }

    private void WriteLegend(LegendPosition position)
    {
        // LEGEND记录 (0x1041)
        WriteRecordHeader(0x1041, 12);

        // 位置
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)position);
        _position += 4;

        // 标志位
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0x0001);
        _position += 4;

        // 预留
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;
    }

    private void WritePlotArea(ChartData chart)
    {
        // PLOTAREA记录 (0x1035)
        WriteRecordHeader(0x1035, 16);

        // 绘图区位置和大小（以1/4000为单位）
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)chart.PlotArea.X);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)chart.PlotArea.Y);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)chart.PlotArea.Width);
        _position += 4;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), (uint)chart.PlotArea.Height);
        _position += 4;

        // 写入高级图表类型的绘图区设置
        WritePlotAreaAdvancedSettings(chart);
    }

    private void WritePlotAreaAdvancedSettings(ChartData chart)
    {
        // 气泡图设置
        if (chart.Type == ChartType.Bubble)
        {
            WriteBubblePlotSettings(chart.PlotArea);
        }

        // 雷达图设置
        if (chart.Type == ChartType.Radar || chart.Type == ChartType.RadarWithMarkers)
        {
            WriteRadarPlotSettings(chart.PlotArea);
        }

        // 股价图设置
        if (chart.PlotArea.StockSettings != null)
        {
            WriteStockPlotSettings(chart.PlotArea.StockSettings);
        }

        // 曲面图设置
        if (chart.PlotArea.SurfaceViewSettings != null)
        {
            WriteSurfacePlotSettings(chart.PlotArea.SurfaceViewSettings);
        }
    }

    private void WriteBubblePlotSettings(ChartPlotArea plotArea)
    {
        // BUBBLEPLOTSETTINGS记录 (0x1048)
        WriteRecordHeader(0x1048, 8);

        // 气泡缩放比例
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)plotArea.BubbleScale);
        _position += 2;

        // 是否显示负值气泡
        _buffer[_position++] = (byte)(plotArea.ShowNegativeBubbles ? 1 : 0);

        // 预留
        _buffer[_position++] = 0;
        _buffer[_position++] = 0;
        _buffer[_position++] = 0;

        // 预留
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;
    }

    private void WriteRadarPlotSettings(ChartPlotArea plotArea)
    {
        // RADARPLOTSETTINGS记录 (0x1049)
        WriteRecordHeader(0x1049, 4);

        // 雷达图样式
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)plotArea.RadarStyle);
        _position += 2;

        // 是否显示轴标签
        _buffer[_position++] = (byte)(plotArea.RadarAxisLabels ? 1 : 0);

        // 预留
        _buffer[_position++] = 0;
    }

    private void WriteStockPlotSettings(StockSettings settings)
    {
        // STOCKPLOTSETTINGS记录 (0x104A)
        WriteRecordHeader(0x104A, 16);

        // 标志位
        var flags = 0u;
        if (settings.ShowDropLines) flags |= 0x0001;
        if (settings.ShowHighLowLines) flags |= 0x0002;
        if (settings.ShowOpenCloseBars) flags |= 0x0004;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 上涨颜色 (RGB)
        var upRgb = (uint)((settings.UpBarColor.R << 16) | (settings.UpBarColor.G << 8) | settings.UpBarColor.B);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), upRgb);
        _position += 4;

        // 下跌颜色 (RGB)
        var downRgb = (uint)((settings.DownBarColor.R << 16) | (settings.DownBarColor.G << 8) | settings.DownBarColor.B);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), downRgb);
        _position += 4;

        // 高低线颜色 (RGB)
        var hlRgb = (uint)((settings.HighLowLineColor.R << 16) | (settings.HighLowLineColor.G << 8) | settings.HighLowLineColor.B);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), hlRgb);
        _position += 4;
    }

    private void WriteSurfacePlotSettings(SurfaceViewSettings settings)
    {
        // SURFACEPLOTSETTINGS记录 (0x104B)
        WriteRecordHeader(0x104B, 20);

        // X轴旋转角度
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)settings.RotationX);
        _position += 2;

        // Y轴旋转角度
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)settings.RotationY);
        _position += 2;

        // 透视角度
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)settings.Perspective);
        _position += 2;

        // 高度百分比
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)settings.HeightPercent);
        _position += 2;

        // 深度百分比
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)settings.DepthPercent);
        _position += 2;

        // 标志位
        var flags = 0u;
        if (settings.RightAngleAxes) flags |= 0x0001;
        if (settings.AutoScaling) flags |= 0x0002;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 墙壁厚度
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)settings.WallThickness);
        _position += 2;

        // 地板厚度
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)settings.FloorThickness);
        _position += 2;
    }

    private void WriteSeries(ChartSeries series, int seriesIndex)
    {
        // SERIES记录 (0x1003)
        WriteRecordHeader(0x1003, 8);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)seriesIndex);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(series.CategoryIndex));
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(series.ValueIndex));
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(series.BubbleIndex));
        _position += 2;

        // 系列名称
        if (!string.IsNullOrEmpty(series.Name))
        {
            WriteSeriesName(series.Name);
        }

        // 类别数据
        if (series.Categories != null)
        {
            WriteCategoryRange(series.Categories);
        }

        // 数值数据
        if (series.Values != null)
        {
            WriteValueRange(series.Values);
        }

        // X值（散点图和气泡图）
        if (series.XValues != null)
        {
            WriteXValueRange(series.XValues);
        }

        // Y值（散点图和气泡图）
        if (series.YValues != null)
        {
            WriteYValueRange(series.YValues);
        }

        // 气泡大小（气泡图）
        if (series.BubbleSizes != null)
        {
            WriteBubbleSizeRange(series.BubbleSizes);
        }

        // 数据标签
        if (series.DataLabels?.Show == true)
        {
            WriteDataLabels(series.DataLabels);
        }

        // 系列样式（线条、标记）
        if (series.LineStyle.HasValue || series.MarkerStyle != MarkerStyle.None)
        {
            WriteSeriesStyle(series);
        }

        // 系列颜色
        if (series.FillColor.HasValue)
        {
            WriteSeriesColor(series.FillColor.Value);
        }

        // 数据点级别设置
        if (series.DataPoints?.Count > 0)
        {
            foreach (var point in series.DataPoints)
            {
                WriteDataPoint(point);
            }
        }

        // 趋势线
        if (series.TrendLines?.Count > 0)
        {
            foreach (var trendLine in series.TrendLines)
            {
                WriteTrendLine(trendLine);
            }
        }

        // 误差线
        if (series.ErrorBars != null)
        {
            WriteErrorBars(series.ErrorBars);
        }

        // 系列结束标记
        WriteRecordHeader(0x1004, 0);
    }

    private void WriteDataPoint(ChartDataPoint point)
    {
        // DATAPOINT记录 (0x1006)
        WriteRecordHeader(0x1006, 12);

        // 数据点索引
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)point.Index);
        _position += 2;

        // 标志位
        var flags = 0u;
        if (point.FillColor.HasValue) flags |= 0x0001;
        if (point.BorderColor.HasValue) flags |= 0x0002;
        if (point.DataLabels?.Show == true) flags |= 0x0004;
        if (point.Explosion.HasValue) flags |= 0x0008;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 爆炸距离（用于饼图）
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(point.Explosion.GetValueOrDefault() ? 25 : 0));
        _position += 2;

        // 预留
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;

        // 写入数据点颜色
        if (point.FillColor.HasValue)
        {
            WriteSeriesColor(point.FillColor.Value);
        }

        // 写入数据点数据标签
        if (point.DataLabels?.Show == true)
        {
            WriteDataLabels(point.DataLabels);
        }
    }

    private void WriteTrendLine(TrendLine trendLine)
    {
        // TRENDLINE记录 (0x1040)
        var nameBytes = string.IsNullOrEmpty(trendLine.Name) ? Array.Empty<byte>() : Encoding.Unicode.GetBytes(trendLine.Name);
        var recLen = 20 + (nameBytes.Length > 0 ? 4 + nameBytes.Length : 0);

        WriteRecordHeader(0x1040, recLen);

        // 趋势线类型
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)trendLine.Type);
        _position += 2;

        // 阶数（多项式）/ 周期（移动平均）
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)(trendLine.Type == TrendLineType.Polynomial ? trendLine.Order : trendLine.Period));
        _position += 2;

        // 前推
        BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), trendLine.Forward.GetValueOrDefault(0));
        _position += 8;

        // 后推
        BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), trendLine.Backward.GetValueOrDefault(0));
        _position += 8;

        // 标志位
        var flags = 0u;
        if (trendLine.DisplayEquation) flags |= 0x0001;
        if (trendLine.DisplayRSquared) flags |= 0x0002;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 趋势线名称
        if (nameBytes.Length > 0)
        {
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)trendLine.Name!.Length);
            _position += 2;
            _buffer[_position++] = 1; // Unicode
            nameBytes.CopyTo(_buffer.Slice(_position));
            _position += nameBytes.Length;
        }

        // 趋势线线条样式
        if (trendLine.LineColor.HasValue || trendLine.LineStyle != LineStyle.Solid)
        {
            WriteRecordHeader(0x100E, 12);
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)trendLine.LineStyle);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 25);
            _position += 2;
            var lineFlags = trendLine.LineColor.HasValue ? 1u : 0u;
            BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), lineFlags);
            _position += 4;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
            _position += 2;
        }
    }

    private void WriteErrorBars(ErrorBars errorBars)
    {
        // ERRORBARS记录 (0x103D)
        WriteRecordHeader(0x103D, 20);

        // 误差线类型
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)errorBars.Type);
        _position += 2;

        // 值类型
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)errorBars.ValueType);
        _position += 2;

        // 值
        BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), errorBars.Value);
        _position += 8;

        // 标志位
        var flags = errorBars.ShowCap ? 1u : 0u;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 预留
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;

        // 误差线线条样式
        if (errorBars.LineColor.HasValue || errorBars.LineStyle != LineStyle.Solid)
        {
            WriteRecordHeader(0x100E, 12);
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)errorBars.LineStyle);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 25);
            _position += 2;
            var lineFlags = errorBars.LineColor.HasValue ? 1u : 0u;
            BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), lineFlags);
            _position += 4;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
            _position += 2;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
            _position += 2;
        }
    }

    private void WriteDataLabels(DataLabels labels)
    {
        // DATALABELS记录 (0x1025)
        WriteRecordHeader(0x1025, 8);

        // 标志位
        var flags = 0u;
        if (labels.Show) flags |= 0x0001;
        if (labels.ShowValue) flags |= 0x0002;
        if (labels.ShowCategory) flags |= 0x0004;
        if (labels.ShowPercentage) flags |= 0x0008;
        if (labels.ShowSeriesName) flags |= 0x0010;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 位置
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)labels.Position);
        _position += 2;

        // 分隔符（0 = 自动）
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteSeriesStyle(ChartSeries series)
    {
        // LINEFORMAT记录 (0x100E) - 线条样式
        if (series.LineStyle.HasValue)
        {
            WriteRecordHeader(0x100E, 12);

            // 线条样式
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)series.LineStyle.Value);
            _position += 2;

            // 线条宽度（以1/20点为单位）
            var width = series.LineStyle == LineStyle.None ? 0 : 25;
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)width);
            _position += 2;

            // 标志位
            var flags = 0u;
            if (series.BorderColor.HasValue) flags |= 0x0001; // 使用自定义颜色
            BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
            _position += 4;

            // 颜色索引（如果使用调色板）
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
            _position += 2;

            // 预留
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
            _position += 2;
        }

        // MARKERFORMAT记录 (0x100F) - 标记样式
        if (series.MarkerStyle != MarkerStyle.None)
        {
            WriteRecordHeader(0x100F, 16);

            // 标记样式
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)series.MarkerStyle);
            _position += 2;

            // 标记大小（以1/20点为单位）
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 60);
            _position += 2;

            // 标志位
            var flags = 0u;
            if (series.FillColor.HasValue) flags |= 0x0001;
            if (series.BorderColor.HasValue) flags |= 0x0002;
            BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
            _position += 4;

            // 填充颜色索引
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
            _position += 2;

            // 边框颜色索引
            BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
            _position += 2;

            // 预留
            BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
            _position += 4;
        }
    }

    private void WriteSeriesColor(ChartColor color)
    {
        // AREAFORMAT记录 (0x100A) - 填充颜色
        WriteRecordHeader(0x100A, 16);

        // 颜色（RGB）
        var rgb = (uint)((color.R << 16) | (color.G << 8) | color.B);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), rgb);
        _position += 4;

        // 背景颜色（透明）
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0xFFFFFFFF);
        _position += 4;

        // 图案（0 = 实心）
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;

        // 标志位
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0001);
        _position += 2;

        // 颜色索引
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;

        // 背景颜色索引
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0xFFFF);
        _position += 2;
    }

    private void WriteSeriesName(string name)
    {
        // SERIESTEXT记录 (0x100D)
        var bytes = Encoding.Unicode.GetBytes(name);
        var recLen = 4 + bytes.Length;

        WriteRecordHeader(0x100D, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)name.Length);
        _position += 2;
        _buffer[_position++] = 1; // Unicode
        bytes.CopyTo(_buffer.Slice(_position));
        _position += bytes.Length;
    }

    private void WriteCategoryRange(ChartRange range)
    {
        // CATEGORY记录 (0x1012)
        WriteRecordHeader(0x1012, 12);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstCol);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastCol);
        _position += 2;

        // 引用类型标志
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0001);
        _position += 2;

        // 工作表索引
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteValueRange(ChartRange range)
    {
        // VALUES记录 (0x1013)
        WriteRecordHeader(0x1013, 12);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstCol);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastCol);
        _position += 2;

        // 引用类型标志
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0001);
        _position += 2;

        // 工作表索引
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteXValueRange(ChartRange range)
    {
        // XVALUES记录 (0x1014) - 用于散点图和气泡图
        WriteRecordHeader(0x1014, 12);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstCol);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastCol);
        _position += 2;

        // 引用类型标志
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0001);
        _position += 2;

        // 工作表索引
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteYValueRange(ChartRange range)
    {
        // YVALUES记录 (0x1015) - 用于散点图和气泡图
        WriteRecordHeader(0x1015, 12);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstCol);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastCol);
        _position += 2;

        // 引用类型标志
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0001);
        _position += 2;

        // 工作表索引
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteBubbleSizeRange(ChartRange range)
    {
        // BUBBLESIZERANGE记录 (0x1047) - 用于气泡图
        WriteRecordHeader(0x1047, 12);

        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstCol);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastCol);
        _position += 2;

        // 引用类型标志
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0x0001);
        _position += 2;

        // 工作表索引
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteAxis(ChartAxis axis)
    {
        // AXIS记录 (0x101D)
        WriteRecordHeader(0x101D, 18);

        // 轴类型
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)axis.Type);
        _position += 2;

        // 轴位置
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)axis.Position);
        _position += 2;

        // 标志位
        var flags = 0u;
        if (axis.HasMajorGridlines) flags |= 0x0001;
        if (axis.HasMinorGridlines) flags |= 0x0002;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 最小值/最大值 (如果是数值轴)
        if (axis.MinValue.HasValue)
        {
            BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), axis.MinValue.Value);
        }
        else
        {
            BinaryPrimitives.WriteUInt64LittleEndian(_buffer.Slice(_position), 0xFFFFFFFFFFFFFFFF);
        }
        _position += 8;

        if (axis.MaxValue.HasValue)
        {
            BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), axis.MaxValue.Value);
        }
        else
        {
            BinaryPrimitives.WriteUInt64LittleEndian(_buffer.Slice(_position), 0xFFFFFFFFFFFFFFFF);
        }
        _position += 8;

        // 轴标题
        if (!string.IsNullOrEmpty(axis.Title))
        {
            WriteAxisTitle(axis.Title);
        }

        // 轴结束标记
        WriteRecordHeader(0x101E, 0);
    }

    private void WriteAxisTitle(string title)
    {
        // AXISTITLE记录 (0x102E)
        var bytes = Encoding.Unicode.GetBytes(title);
        var recLen = 4 + bytes.Length;

        WriteRecordHeader(0x102E, recLen);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)title.Length);
        _position += 2;
        _buffer[_position++] = 1; // Unicode
        bytes.CopyTo(_buffer.Slice(_position));
        _position += bytes.Length;
    }

    private void WriteAxisLink()
    {
        // AXISLINK记录 (0x1026)
        WriteRecordHeader(0x1026, 2);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;
    }

    private void WriteDataTable(ChartDataTable dataTable)
    {
        // DATATABLE记录 (0x1036)
        WriteRecordHeader(0x1036, 12);

        // 标志位
        var flags = 0u;
        if (dataTable.ShowLegendKeys) flags |= 0x0001;
        if (dataTable.HasHorizontalBorder) flags |= 0x0002;
        if (dataTable.HasVerticalBorder) flags |= 0x0004;
        if (dataTable.HasOutlineBorder) flags |= 0x0008;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 字体大小（以1/20点为单位）
        var fontSize = (ushort)(dataTable.FontSize * 20);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), fontSize);
        _position += 2;

        // 预留字段
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;

        // 数据表结束标记
        WriteRecordHeader(0x1037, 0);
    }

    private void WriteEof()
    {
        WriteRecordHeader(0x000A, 0);
    }

    private void WriteRecordHeader(ushort type, int length)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), type);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)length);
        _position += 2;
    }
}
