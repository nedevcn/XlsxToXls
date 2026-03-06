using System;
using System.Buffers.Binary;

namespace Nedev.FileConverters.XlsxToXls.Internal
{
    internal static class BufferHelpers
    {
        public static void WriteDoubleLittleEndian(Span<byte> span, double value)
        {
#if NETSTANDARD2_1
            // netstandard2.1 does not include WriteDoubleLittleEndian
            var bytes = BitConverter.GetBytes(value);
            if (!BitConverter.IsLittleEndian)
                Array.Reverse(bytes);
            bytes.CopyTo(span);
#else
            BinaryPrimitives.WriteDoubleLittleEndian(span, value);
#endif
        }
    }
}
