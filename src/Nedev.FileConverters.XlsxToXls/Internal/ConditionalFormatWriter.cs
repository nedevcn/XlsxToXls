using System.Buffers;
using System.Buffers.Binary;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// BIFF8条件格式记录写入器 - 使用ArrayPool减少内存分配
/// </summary>
internal ref struct ConditionalFormatWriter
{
    private Span<byte> _buffer;
    private int _position;
    private byte[]? _pooledBuffer;

    public ConditionalFormatWriter(Span<byte> buffer)
    {
        _buffer = buffer;
        _position = 0;
        _pooledBuffer = null;
    }

    /// <summary>
    /// 使用ArrayPool创建ConditionalFormatWriter，自动管理缓冲区
    /// </summary>
    public static ConditionalFormatWriter CreatePooled(out byte[] pooledBuffer, int minSize = 16384)
    {
        pooledBuffer = ArrayPool<byte>.Shared.Rent(minSize);
        return new ConditionalFormatWriter(pooledBuffer.AsSpan())
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
    /// 写入条件格式规则
    /// </summary>
    public int WriteConditionalFormat(ConditionalFormatData format, int sheetIndex)
    {
        // CFHEADER记录 (0x01B0)
        WriteCfHeader(format);

        // CFRULE记录 (0x01B1)
        WriteCfRule(format);

        // 写入范围
        foreach (var range in format.Ranges)
        {
            WriteCfRange(range);
        }

        // 写入样式/格式
        if (format.Style != null)
        {
            WriteCfStyle(format.Style);
        }

        // 写入色阶
        if (format.ColorScale != null)
        {
            WriteColorScale(format.ColorScale);
        }

        // 写入数据条
        if (format.DataBar != null)
        {
            WriteDataBar(format.DataBar);
        }

        // 写入图标集
        if (format.IconSet != null)
        {
            WriteIconSet(format.IconSet);
        }

        return _position;
    }

    private void WriteCfHeader(ConditionalFormatData format)
    {
        // CFHEADER记录 (0x01B0)
        var recLen = 8 + (format.Ranges.Count * 8);
        WriteRecordHeader(0x01B0, recLen);

        // 规则数量
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 1);
        _position += 2;

        // 单元格范围数量
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)format.Ranges.Count);
        _position += 2;

        // 标志位
        var flags = format.StopIfTrue ? 0x0001u : 0u;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 写入每个范围
        foreach (var range in format.Ranges)
        {
            WriteRangeCoords(range);
        }
    }

    private void WriteRangeCoords(CellRange range)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastRow);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.FirstCol);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)range.LastCol);
        _position += 2;
    }

    private void WriteCfRule(ConditionalFormatData format)
    {
        // CFRULE记录 (0x01B1)
        var recLen = 12; // 基础大小

        // 计算公式长度
        var formula1Len = string.IsNullOrEmpty(format.Formula1) ? 0 : GetFormulaTokenLength(format.Formula1);
        var formula2Len = string.IsNullOrEmpty(format.Formula2) ? 0 : GetFormulaTokenLength(format.Formula2);

        recLen += formula1Len + formula2Len;

        WriteRecordHeader(0x01B1, recLen);

        // 规则类型
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)format.Type);
        _position += 2;

        // 操作符
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)format.Operator);
        _position += 2;

        // 标志位
        var flags = 0u;
        if (format.StopIfTrue) flags |= 0x0001;
        if (!string.IsNullOrEmpty(format.Formula1)) flags |= 0x0002;
        if (!string.IsNullOrEmpty(format.Formula2)) flags |= 0x0004;
        if (format.Style != null) flags |= 0x0008;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 公式1
        if (!string.IsNullOrEmpty(format.Formula1))
        {
            WriteFormulaTokens(format.Formula1);
        }

        // 公式2
        if (!string.IsNullOrEmpty(format.Formula2))
        {
            WriteFormulaTokens(format.Formula2);
        }
    }

    private void WriteCfRange(CellRange range)
    {
        // 范围已经包含在CFHEADER中
        // 这里可以添加额外的范围特定记录
    }

    private void WriteCfStyle(ConditionalFormatStyle style)
    {
        // CFSTYLE记录 (0x01B2)
        var recLen = 20;
        WriteRecordHeader(0x01B2, recLen);

        // 标志位
        var flags = 0u;
        if (style.FontColor.HasValue) flags |= 0x0001;
        if (style.Bold.HasValue) flags |= 0x0002;
        if (style.Italic.HasValue) flags |= 0x0004;
        if (style.FillColor.HasValue) flags |= 0x0008;
        if (style.Border != null) flags |= 0x0010;
        if (style.NumberFormat != null) flags |= 0x0020;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 字体颜色
        if (style.FontColor.HasValue)
        {
            WriteRgbColor(style.FontColor.Value);
        }
        else
        {
            BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
            _position += 4;
        }

        // 填充颜色
        if (style.FillColor.HasValue)
        {
            WriteRgbColor(style.FillColor.Value);
        }
        else
        {
            BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
            _position += 4;
        }

        // 字体样式
        var fontFlags = 0u;
        if (style.Bold == true) fontFlags |= 0x0001;
        if (style.Italic == true) fontFlags |= 0x0002;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), fontFlags);
        _position += 4;

        // 边框（简化处理）
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), 0);
        _position += 4;
    }

    private void WriteColorScale(ColorScale colorScale)
    {
        // COLORSCALE记录 (0x01B3)
        var recLen = 4 + (colorScale.Midpoint != null ? 24 : 16);
        WriteRecordHeader(0x01B3, recLen);

        // 颜色点数量
        var pointCount = colorScale.Midpoint != null ? (ushort)3 : (ushort)2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), pointCount);
        _position += 2;

        // 标志位
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), 0);
        _position += 2;

        // 最小值点
        WriteColorScalePoint(colorScale.Minimum);

        // 中点值（如果有）
        if (colorScale.Midpoint != null)
        {
            WriteColorScalePoint(colorScale.Midpoint);
        }

        // 最大值点
        WriteColorScalePoint(colorScale.Maximum);
    }

    private void WriteColorScalePoint(ColorScalePoint point)
    {
        // 值类型
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)point.Type);
        _position += 2;

        // 值（如果是数值类型）
        if (point.Value.HasValue)
        {
            BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), point.Value.Value);
        }
        else
        {
            BinaryPrimitives.WriteUInt64LittleEndian(_buffer.Slice(_position), 0);
        }
        _position += 8;

        // 颜色
        WriteRgbColor(point.Color);
    }

    private void WriteDataBar(DataBar dataBar)
    {
        // DATABAR记录 (0x01B4)
        var recLen = 24;
        WriteRecordHeader(0x01B4, recLen);

        // 标志位
        var flags = 0u;
        if (dataBar.ShowValue) flags |= 0x0001;
        if (dataBar.BorderColor.HasValue) flags |= 0x0002;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 最小值类型和值
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)dataBar.Minimum.Type);
        _position += 2;
        if (dataBar.Minimum.Value.HasValue)
        {
            BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), dataBar.Minimum.Value.Value);
        }
        else
        {
            BinaryPrimitives.WriteUInt64LittleEndian(_buffer.Slice(_position), 0);
        }
        _position += 8;

        // 最大值类型和值
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)dataBar.Maximum.Type);
        _position += 2;
        if (dataBar.Maximum.Value.HasValue)
        {
            BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), dataBar.Maximum.Value.Value);
        }
        else
        {
            BinaryPrimitives.WriteUInt64LittleEndian(_buffer.Slice(_position), 0);
        }
        _position += 8;

        // 颜色
        WriteRgbColor(dataBar.Color);
    }

    private void WriteIconSet(IconSet iconSet)
    {
        // ICONSET记录 (0x01B5)
        var recLen = 8 + (iconSet.Thresholds?.Count * 10 ?? 0);
        WriteRecordHeader(0x01B5, recLen);

        // 图标集类型
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)iconSet.Type);
        _position += 2;

        // 标志位
        var flags = 0u;
        if (iconSet.ShowValue) flags |= 0x0001;
        if (iconSet.Reverse) flags |= 0x0002;
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), flags);
        _position += 4;

        // 阈值数量
        var thresholdCount = (ushort)(iconSet.Thresholds?.Count ?? 0);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), thresholdCount);
        _position += 2;

        // 写入阈值
        if (iconSet.Thresholds != null)
        {
            foreach (var threshold in iconSet.Thresholds)
            {
                WriteIconThreshold(threshold);
            }
        }
    }

    private void WriteIconThreshold(IconThreshold threshold)
    {
        // 值类型
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)threshold.Type);
        _position += 2;

        // 值
        BufferHelpers.WriteDoubleLittleEndian(_buffer.Slice(_position), threshold.Value);
        _position += 8;
    }

    private void WriteRgbColor(ChartColor color)
    {
        var rgb = (uint)((color.R << 16) | (color.G << 8) | color.B);
        BinaryPrimitives.WriteUInt32LittleEndian(_buffer.Slice(_position), rgb);
        _position += 4;
    }

    private int GetFormulaTokenLength(string formula)
    {
        // 简化处理：返回估算的公式令牌长度
        // 实际实现需要完整的公式解析器
        return 16 + (formula.Length * 2);
    }

    private void WriteFormulaTokens(string formula)
    {
        // 简化处理：写入公式字符串
        // 实际实现需要将公式编译为BIFF8令牌
        var len = (ushort)Math.Min(formula.Length, 255);
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), len);
        _position += 2;

        // 写入公式字符（简化）
        for (var i = 0; i < len && i < formula.Length; i++)
        {
            _buffer[_position++] = (byte)formula[i];
        }
    }

    private void WriteRecordHeader(ushort type, int length)
    {
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), type);
        _position += 2;
        BinaryPrimitives.WriteUInt16LittleEndian(_buffer.Slice(_position), (ushort)length);
        _position += 2;
    }
}
