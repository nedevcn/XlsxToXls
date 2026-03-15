using System.Buffers;
using System.Runtime.CompilerServices;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// High-performance memory pool manager for reducing GC pressure and memory allocations.
/// Wraps <see cref="ArrayPool{T}"/> with convenience methods and predefined buffer sizes.
/// </summary>
/// <remarks>
/// <para><b>Example - Basic usage:</b></para>
/// <code language="csharp">
/// // Rent a buffer
/// var buffer = MemoryPool.Rent(1024);
/// try
/// {
///     // Use the buffer
///     var span = buffer.AsSpan(0, 1024);
///     // ... write data to span ...
/// }
/// finally
/// {
///     // Always return the buffer to the pool
///     MemoryPool.Return(buffer);
/// }
/// </code>
/// <para><b>Example - Using PooledBuffer struct:</b></para>
/// <code language="csharp">
/// using (var pooled = new PooledBuffer(4096))
/// {
///     var span = pooled.Span;
///     // ... use span ...
/// } // Buffer automatically returned to pool
/// </code>
/// </remarks>
internal static class MemoryPool
{
    // 常用缓冲区大小，避免重复分配
    private const int SmallBufferSize = 1024;
    private const int MediumBufferSize = 4096;
    private const int LargeBufferSize = 16384;
    private const int XLargeBufferSize = 65536;

    /// <summary>
    /// 获取适当大小的缓冲区
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static byte[] Rent(int minimumLength)
    {
        return ArrayPool<byte>.Shared.Rent(minimumLength);
    }

    /// <summary>
    /// 返回缓冲区到池中
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static void Return(byte[] buffer, bool clearArray = false)
    {
        if (buffer != null)
        {
            ArrayPool<byte>.Shared.Return(buffer, clearArray);
        }
    }

    /// <summary>
    /// 获取小型缓冲区 (1KB)
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static byte[] RentSmall() => Rent(SmallBufferSize);

    /// <summary>
    /// 获取中型缓冲区 (4KB)
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static byte[] RentMedium() => Rent(MediumBufferSize);

    /// <summary>
    /// 获取大型缓冲区 (16KB)
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static byte[] RentLarge() => Rent(LargeBufferSize);

    /// <summary>
    /// 获取超大缓冲区 (64KB)
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static byte[] RentXLarge() => Rent(XLargeBufferSize);

    /// <summary>
    /// 计算实际需要的最小缓冲区大小（向上取整到2的幂）
    /// </summary>
    public static int CalculateOptimalSize(int requiredSize)
    {
        if (requiredSize <= SmallBufferSize) return SmallBufferSize;
        if (requiredSize <= MediumBufferSize) return MediumBufferSize;
        if (requiredSize <= LargeBufferSize) return LargeBufferSize;
        if (requiredSize <= XLargeBufferSize) return XLargeBufferSize;

        // 对于更大的缓冲区，向上取整到64KB的倍数
        return ((requiredSize + XLargeBufferSize - 1) / XLargeBufferSize) * XLargeBufferSize;
    }
}

/// <summary>
/// 可租用的内存缓冲区包装器 - 支持using语句
/// </summary>
internal readonly struct PooledBuffer : IDisposable
{
    private readonly byte[] _buffer;
    private readonly bool _clearOnReturn;

    public PooledBuffer(int minimumLength, bool clearOnReturn = false)
    {
        _buffer = MemoryPool.Rent(minimumLength);
        _clearOnReturn = clearOnReturn;
    }

    public Span<byte> Span => _buffer;
    public Memory<byte> Memory => _buffer.AsMemory();
    public byte[] Array => _buffer;
    public int Length => _buffer.Length;

    public void Dispose()
    {
        MemoryPool.Return(_buffer, _clearOnReturn);
    }
}

/// <summary>
/// StringBuilder池 - 重用StringBuilder实例
/// </summary>
internal static class StringBuilderPool
{
    [ThreadStatic]
    private static System.Text.StringBuilder? _cachedInstance;

    /// <summary>
    /// 获取StringBuilder实例
    /// </summary>
    public static System.Text.StringBuilder Rent(int capacity = 256)
    {
        var sb = _cachedInstance;
        if (sb != null)
        {
            _cachedInstance = null;
            sb.Clear();
            if (sb.Capacity < capacity)
            {
                sb.Capacity = capacity;
            }
            return sb;
        }
        return new System.Text.StringBuilder(capacity);
    }

    /// <summary>
    /// 返回StringBuilder实例到池中
    /// </summary>
    public static void Return(System.Text.StringBuilder sb)
    {
        if (sb != null && sb.Capacity <= 4096) // 只缓存较小的实例
        {
            _cachedInstance = sb;
        }
    }

    /// <summary>
    /// 获取StringBuilder，使用完成后自动返回池中
    /// </summary>
    public static string ToStringAndReturn(System.Text.StringBuilder sb)
    {
        var result = sb.ToString();
        Return(sb);
        return result;
    }
}

/// <summary>
/// 列表池 - 重用List实例
/// </summary>
internal static class ListPool<T>
{
    [ThreadStatic]
    private static List<T>? _cachedInstance;

    /// <summary>
    /// 获取List实例
    /// </summary>
    public static List<T> Rent(int capacity = 16)
    {
        var list = _cachedInstance;
        if (list != null)
        {
            _cachedInstance = null;
            list.Clear();
            if (list.Capacity < capacity)
            {
                list.Capacity = capacity;
            }
            return list;
        }
        return new List<T>(capacity);
    }

    /// <summary>
    /// 返回List实例到池中
    /// </summary>
    public static void Return(List<T> list)
    {
        if (list != null && list.Capacity <= 1024) // 只缓存较小的实例
        {
            list.Clear();
            _cachedInstance = list;
        }
    }
}
