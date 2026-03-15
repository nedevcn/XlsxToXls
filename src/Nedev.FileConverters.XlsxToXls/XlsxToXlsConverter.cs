using System.Buffers;
using System.Buffers.Binary;
using System.Text;
using Nedev.FileConverters.XlsxToXls.Internal;
using Nedev.FileConverters.Core;

namespace Nedev.FileConverters.XlsxToXls;

using Nedev.FileConverters.Core;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using System.Linq;

/// <summary>
/// Initializes the XLSX to XLS converter.
/// </summary>
public static class XlsxToXlsConverterInitializer
{
    private static int _initialized;
    private static readonly object _lock = new();

    /// <summary>
    /// Gets the lock object used for thread-safe initialization.
    /// </summary>
    internal static object LockObject => _lock;

    /// <summary>
    /// Registers the CodePages encoding provider required for BIFF8 format.
    /// This method is called automatically before conversion.
    /// Thread-safe for concurrent access.
    /// </summary>
    public static void EnsureInitialized()
    {
        if (Interlocked.CompareExchange(ref _initialized, 1, 0) == 0)
        {
            lock (_lock)
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            }
        }
    }
}

/// <summary>
/// High-performance XLSX to XLS converter with zero third-party dependencies.
/// Uses streaming and pooled buffers for optimal memory usage.
/// </summary>
/// <remarks>
/// <para>
/// This converter supports:
/// <list type="bullet">
///   <item><description>All Excel cell data types (numbers, text, dates, booleans)</description></item>
///   <item><description>Formulas with 20+ built-in functions</description></item>
///   <item><description>Cell formatting and styles</description></item>
///   <item><description>Charts (8 types with data labels, trendlines, error bars)</description></item>
///   <item><description>Merge cells, hyperlinks, data validation</description></item>
///   <item><description>Comments and conditional formatting</description></item>
/// </list>
/// </para>
/// <para><b>Example usage:</b></para>
/// <code language="csharp">
/// // Basic conversion
/// using (var input = File.OpenRead("input.xlsx"))
/// using (var output = File.Create("output.xls"))
/// {
///     XlsxToXlsConverter.Convert(input, output);
/// }
/// 
/// // Conversion with logging
/// using (var input = File.OpenRead("input.xlsx"))
/// using (var output = File.Create("output.xls"))
/// {
///     XlsxToXlsConverter.Convert(input, output, log => Console.WriteLine(log));
/// }
/// 
/// // Using the Core framework adapter
/// var adapter = new XlsxToXlsConverter.FileConverterAdapter();
/// using (var input = File.OpenRead("input.xlsx"))
/// using (var result = adapter.Convert(input))
/// {
///     using (var output = File.Create("output.xls"))
///     {
///         result.CopyTo(output);
///     }
/// }
/// </code>
/// </remarks>
public static class XlsxToXlsConverter
{
    /// <summary>
    /// Adapter for Nedev.FileConverters.Core integration.
    /// Enables this converter to be used within the core framework.
    /// </summary>
    [FileConverter("xlsx", "xls")]
    public class FileConverterAdapter : IFileConverter
    {
        /// <summary>
        /// Converts the input XLSX stream to XLS format.
        /// </summary>
        /// <param name="input">The input XLSX stream to convert.</param>
        /// <returns>A new stream containing the converted XLS data. The caller is responsible for disposing this stream.</returns>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="input"/> is null.</exception>
        /// <exception cref="InvalidDataException">Thrown when the input is not a valid XLSX file.</exception>
        public Stream Convert(Stream input)
        {
            if (input == null)
                throw new ArgumentNullException(nameof(input));
            var output = new MemoryStream();
            XlsxToXlsConverter.Convert(input, output);
            output.Position = 0;
            return output;
        }
    }

    /// <summary>
    /// Converts XLSX stream to XLS format and writes to output stream.
    /// </summary>
    /// <param name="xlsxStream">Input XLSX stream. Should be readable and seekable for optimal performance.</param>
    /// <param name="xlsStream">Output XLS stream. Must be writable.</param>
    /// <param name="log">Optional logging callback for diagnostic messages. Called when formulas or other features cannot be fully converted.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="xlsxStream"/> or <paramref name="xlsStream"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="xlsxStream"/> is not readable or <paramref name="xlsStream"/> is not writable.</exception>
    /// <exception cref="InvalidDataException">Thrown when the input is not a valid XLSX file.</exception>
    /// <exception cref="IOException">Thrown when an I/O error occurs during conversion.</exception>
    /// <remarks>
    /// The conversion process:
    /// <list type="number">
    ///   <item><description>Reads XLSX structure (shared strings, sheets, styles)</description></item>
    ///   <item><description>Estimates required BIFF8 buffer size</description></item>
    ///   <item><description>Writes BIFF8 records (workbook, sheets, cells, formatting)</description></item>
    ///   <item><description>Wraps in OLE Compound File format</description></item>
    /// </list>
    /// </remarks>
    public static void Convert(Stream xlsxStream, Stream xlsStream, Action<string>? log = null)
    {
        XlsxToXlsConverterInitializer.EnsureInitialized();

        if (xlsxStream == null)
            throw new ArgumentNullException(nameof(xlsxStream));
        if (xlsStream == null)
            throw new ArgumentNullException(nameof(xlsStream));

        if (!xlsxStream.CanRead)
            throw new ArgumentException("Input stream must be readable.", nameof(xlsxStream));
        if (!xlsStream.CanWrite)
            throw new ArgumentException("Output stream must be writable.", nameof(xlsStream));

        log?.Invoke($"[XlsxToXlsConverter] Starting conversion...");

        var (sharedStrings, sheets, styles, definedNames, docSecurity) = XlsxReader.Read(xlsxStream, log);
        log?.Invoke($"[XlsxToXlsConverter] Read {sharedStrings.Count} shared strings, {sheets.Count} sheets, {definedNames.Count} defined names");
        if (docSecurity.HasValue && docSecurity.Value.IsSigned)
        {
            var sigCount = docSecurity.Value.Signatures?.Count ?? 0;
            log?.Invoke($"[XlsxToXlsConverter] Document has {sigCount} digital signature(s)");
        }

        var stylesData = styles ?? CreateDefaultStyles();
        log?.Invoke($"[XlsxToXlsConverter] Using {stylesData.Fonts.Count} fonts, {stylesData.NumFmts.Count} number formats, {stylesData.CellXfs.Count} cell styles");

        var biffSize = EstimateBiffSize(sharedStrings, sheets, stylesData, definedNames);
        log?.Invoke($"[XlsxToXlsConverter] Estimated BIFF size: {biffSize} bytes");

        var buffer = ArrayPool<byte>.Shared.Rent(Math.Max(biffSize, 256 * 1024));
        try
        {
            var written = WriteBiff(buffer.AsSpan(), sharedStrings, sheets, stylesData, definedNames, log);
            log?.Invoke($"[XlsxToXlsConverter] Written {written} bytes of BIFF data");

            var ole = new OleCompoundWriter("Workbook");
            ole.WriteStream(buffer.AsSpan(0, written));
            ole.WriteTo(xlsStream);
            log?.Invoke($"[XlsxToXlsConverter] OLE compound file written successfully");
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(buffer);
            log?.Invoke($"[XlsxToXlsConverter] Conversion completed");
        }
    }

    /// <summary>
    /// Converts XLSX file to XLS file.
    /// </summary>
    /// <param name="xlsxPath">Path to the input XLSX file.</param>
    /// <param name="xlsPath">Path to the output XLS file. Will be created or overwritten.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="xlsxPath"/> or <paramref name="xlsPath"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="xlsxPath"/> is empty or contains only whitespace.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    /// <exception cref="UnauthorizedAccessException">Thrown when access to the file is denied.</exception>
    /// <exception cref="IOException">Thrown when an I/O error occurs during file operations.</exception>
    /// <exception cref="InvalidDataException">Thrown when the input file is not a valid XLSX file.</exception>
    /// <remarks>
    /// This is a convenience method that opens the files and calls <see cref="Convert(Stream, Stream, Action{string}?)"/>.
    /// The files are properly closed after conversion, even if an error occurs.
    /// </remarks>
    public static void ConvertFile(string xlsxPath, string xlsPath)
    {
        if (string.IsNullOrWhiteSpace(xlsxPath))
            throw new ArgumentException("Path cannot be null or whitespace.", nameof(xlsxPath));
        if (string.IsNullOrWhiteSpace(xlsPath))
            throw new ArgumentException("Path cannot be null or whitespace.", nameof(xlsPath));

        using var xlsx = File.OpenRead(xlsxPath);
        using var xls = File.Create(xlsPath);
        Convert(xlsx, xls);
    }

    /// <summary>
    /// Converts XLSX file to XLS file with logging.
    /// </summary>
    /// <param name="xlsxPath">Path to the input XLSX file.</param>
    /// <param name="xlsPath">Path to the output XLS file. Will be created or overwritten.</param>
    /// <param name="log">Logging callback for diagnostic messages.</param>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="xlsxPath"/>, <paramref name="xlsPath"/>, or <paramref name="log"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="xlsxPath"/> is empty or contains only whitespace.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    /// <exception cref="UnauthorizedAccessException">Thrown when access to the file is denied.</exception>
    /// <exception cref="IOException">Thrown when an I/O error occurs during file operations.</exception>
    /// <exception cref="InvalidDataException">Thrown when the input file is not a valid XLSX file.</exception>
    /// <remarks>
    /// This overload allows you to receive diagnostic messages during conversion.
    /// Log messages may include:
    /// <list type="bullet">
    ///   <item><description>Unsupported formula functions</description></item>
    ///   <item><description>Chart conversion warnings</description></item>
    ///   <item><description>Style conversion issues</description></item>
    /// </list>
    /// </remarks>
    public static void ConvertFile(string xlsxPath, string xlsPath, Action<string> log)
    {
        if (string.IsNullOrWhiteSpace(xlsxPath))
            throw new ArgumentException("Path cannot be null or whitespace.", nameof(xlsxPath));
        if (string.IsNullOrWhiteSpace(xlsPath))
            throw new ArgumentException("Path cannot be null or whitespace.", nameof(xlsPath));
        if (log == null)
            throw new ArgumentNullException(nameof(log));

        using var xlsx = File.OpenRead(xlsxPath);
        using var xls = File.Create(xlsPath);
        Convert(xlsx, xls, log);
    }

    /// <summary>
    /// Converts XLSX file to XLS file asynchronously.
    /// </summary>
    /// <param name="xlsxPath">Path to the input XLSX file.</param>
    /// <param name="xlsPath">Path to the output XLS file. Will be created or overwritten.</param>
    /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
    /// <returns>A task representing the asynchronous conversion operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="xlsxPath"/> or <paramref name="xlsPath"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="xlsxPath"/> is empty or contains only whitespace.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    /// <exception cref="UnauthorizedAccessException">Thrown when access to the file is denied.</exception>
    /// <exception cref="IOException">Thrown when an I/O error occurs during file operations.</exception>
    /// <exception cref="InvalidDataException">Thrown when the input file is not a valid XLSX file.</exception>
    /// <exception cref="OperationCanceledException">Thrown when the operation is canceled.</exception>
    /// <remarks>
    /// This method performs the conversion asynchronously, allowing the calling thread to continue execution.
    /// The actual conversion is still synchronous internally, but file I/O is performed asynchronously.
    /// </remarks>
    public static async Task ConvertFileAsync(string xlsxPath, string xlsPath, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(xlsxPath))
            throw new ArgumentException("Path cannot be null or whitespace.", nameof(xlsxPath));
        if (string.IsNullOrWhiteSpace(xlsPath))
            throw new ArgumentException("Path cannot be null or whitespace.", nameof(xlsPath));

        cancellationToken.ThrowIfCancellationRequested();

        using var xlsx = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous);
        using var xls = new FileStream(xlsPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, FileOptions.Asynchronous);
        
        // Read input into memory stream for synchronous processing
        using var memStream = new MemoryStream();
        await xlsx.CopyToAsync(memStream, cancellationToken);
        memStream.Position = 0;
        
        using var outputStream = new MemoryStream();
        Convert(memStream, outputStream);
        
        outputStream.Position = 0;
        await outputStream.CopyToAsync(xls, cancellationToken);
    }

    /// <summary>
    /// Converts XLSX file to XLS file asynchronously with logging.
    /// </summary>
    /// <param name="xlsxPath">Path to the input XLSX file.</param>
    /// <param name="xlsPath">Path to the output XLS file. Will be created or overwritten.</param>
    /// <param name="log">Logging callback for diagnostic messages.</param>
    /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
    /// <returns>A task representing the asynchronous conversion operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="xlsxPath"/>, <paramref name="xlsPath"/>, or <paramref name="log"/> is null.</exception>
    /// <exception cref="ArgumentException">Thrown when <paramref name="xlsxPath"/> is empty or contains only whitespace.</exception>
    /// <exception cref="FileNotFoundException">Thrown when the input file does not exist.</exception>
    /// <exception cref="UnauthorizedAccessException">Thrown when access to the file is denied.</exception>
    /// <exception cref="IOException">Thrown when an I/O error occurs during file operations.</exception>
    /// <exception cref="InvalidDataException">Thrown when the input file is not a valid XLSX file.</exception>
    /// <exception cref="OperationCanceledException">Thrown when the operation is canceled.</exception>
    public static async Task ConvertFileAsync(string xlsxPath, string xlsPath, Action<string> log, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(xlsxPath))
            throw new ArgumentException("Path cannot be null or whitespace.", nameof(xlsxPath));
        if (string.IsNullOrWhiteSpace(xlsPath))
            throw new ArgumentException("Path cannot be null or whitespace.", nameof(xlsPath));
        if (log == null)
            throw new ArgumentNullException(nameof(log));

        cancellationToken.ThrowIfCancellationRequested();

        using var xlsx = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous);
        using var xls = new FileStream(xlsPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, FileOptions.Asynchronous);
        
        using var memStream = new MemoryStream();
        await xlsx.CopyToAsync(memStream, cancellationToken);
        memStream.Position = 0;
        
        using var outputStream = new MemoryStream();
        Convert(memStream, outputStream, log);
        
        outputStream.Position = 0;
        await outputStream.CopyToAsync(xls, cancellationToken);
    }

    /// <summary>
    /// Converts multiple XLSX files to XLS format in batch.
    /// </summary>
    /// <param name="files">Collection of tuples containing input XLSX path and output XLS path.</param>
    /// <param name="parallelOptions">Options for parallel execution, including degree of parallelism and cancellation token.</param>
    /// <returns>A task representing the batch conversion operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="files"/> is null.</exception>
    /// <exception cref="OperationCanceledException">Thrown when the operation is canceled.</exception>
    /// <remarks>
    /// <para>
    /// This method converts multiple files sequentially. For parallel conversion, use the overload with progress reporting.
    /// </para>
    /// <para><b>Example usage:</b></para>
    /// <code language="csharp">
    /// var files = new[]
    /// {
    ///     ("input1.xlsx", "output1.xls"),
    ///     ("input2.xlsx", "output2.xls"),
    ///     ("input3.xlsx", "output3.xls")
    /// };
    /// 
    /// await XlsxToXlsConverter.ConvertBatchAsync(files);
    /// </code>
    /// </remarks>
    public static async Task ConvertBatchAsync(IEnumerable<(string xlsxPath, string xlsPath)> files, ParallelOptions? parallelOptions = null)
    {
        if (files == null)
            throw new ArgumentNullException(nameof(files));

        var options = parallelOptions ?? new ParallelOptions { MaxDegreeOfParallelism = 1 };
        var fileList = files.ToList();
        
        await Task.Run(() =>
        {
            foreach (var filePair in fileList)
            {
                options.CancellationToken.ThrowIfCancellationRequested();
                ConvertFile(filePair.xlsxPath, filePair.xlsPath);
            }
        }, options.CancellationToken);
    }

    /// <summary>
    /// Converts multiple XLSX files to XLS format in batch with progress reporting.
    /// </summary>
    /// <param name="files">Collection of tuples containing input XLSX path and output XLS path.</param>
    /// <param name="progress">Progress reporter that receives completion percentage (0-100).</param>
    /// <param name="parallelOptions">Options for parallel execution, including degree of parallelism and cancellation token.</param>
    /// <returns>A task representing the batch conversion operation.</returns>
    /// <exception cref="ArgumentNullException">Thrown when <paramref name="files"/> or <paramref name="progress"/> is null.</exception>
    /// <exception cref="OperationCanceledException">Thrown when the operation is canceled.</exception>
    /// <remarks>
    /// <para>
    /// This method provides progress reporting during batch conversion.
    /// The progress callback is invoked after each file is converted with the current completion percentage.
    /// Files are processed sequentially to ensure thread safety.
    /// </para>
    /// <para><b>Example usage:</b></para>
    /// <code language="csharp">
    /// var files = Directory.GetFiles(".", "*.xlsx")
    ///     .Select(f => (f, Path.ChangeExtension(f, ".xls")));
    /// 
    /// var progress = new Progress&lt;int&gt;(percent => 
    ///     Console.WriteLine($"Progress: {percent}%"));
    /// 
    /// await XlsxToXlsConverter.ConvertBatchAsync(files, progress);
    /// </code>
    /// </remarks>
    public static async Task ConvertBatchAsync(IEnumerable<(string xlsxPath, string xlsPath)> files, IProgress<int> progress, ParallelOptions? parallelOptions = null)
    {
        if (files == null)
            throw new ArgumentNullException(nameof(files));
        if (progress == null)
            throw new ArgumentNullException(nameof(progress));

        var options = parallelOptions ?? new ParallelOptions { MaxDegreeOfParallelism = 1 };
        var fileList = files.ToList();
        var totalFiles = fileList.Count;
        var completedFiles = 0;

        await Task.Run(() =>
        {
            foreach (var filePair in fileList)
            {
                options.CancellationToken.ThrowIfCancellationRequested();
                ConvertFile(filePair.xlsxPath, filePair.xlsPath);
                
                completedFiles++;
                var percent = (int)((double)completedFiles / totalFiles * 100);
                progress.Report(percent);
            }
        }, options.CancellationToken);
    }

    private static StylesData CreateDefaultStyles()
    {
        var s = new StylesData();
        s.EnsureMinFonts();
        s.CellXfs.Add(new CellXfInfo(
            NumFmtId: 0,
            FontId: 0,
            FillId: 0,
            BorderId: 0,
            HorizontalAlign: 0,
            VerticalAlign: 2, // 默认底部对齐
            WrapText: false,
            Indent: 0,
            Locked: true,
            Hidden: false));
        return s;
    }

    private static int EstimateBiffSize(List<ReadOnlyMemory<char>> sharedStrings, List<SheetData> sheets, StylesData styles, List<DefinedNameInfo> definedNames)
    {
        var n = 1024;
        foreach (var s in sharedStrings)
            n += 16 + s.Length * 2;
        foreach (var sheet in sheets)
        {
            n += 64 * sheet.ColInfos.Count + 32 * sheet.MergeRanges.Count;
            foreach (var hl in sheet.Hyperlinks)
                n += 60 + hl.Url.Length * 2;
            foreach (var comment in sheet.Comments)
            {
                n += 64 + 25 + comment.Author.Length * 2;
                if (!string.IsNullOrEmpty(comment.Text))
                    n += 22 + 5 + comment.Text.Length * 2 + 20;
            }
            foreach (var dv in sheet.DataValidations)
                n += 30 + dv.Ranges.Count * 8 + (dv.PromptTitle.Length + dv.ErrorTitle.Length + dv.PromptText.Length + dv.ErrorText.Length) * 2 + dv.Formula1.Length + dv.Formula2.Length;
            // 图表大小估算
            foreach (var chart in sheet.Charts)
            {
                n += 1024 + chart.Series.Count * 256;
                if (!string.IsNullOrEmpty(chart.Title?.Text))
                    n += 64 + chart.Title.Text.Length * 2;
            }
        }
        if (sheets.Count > 0)
            n += 8 + (4 + 2 + sheets.Count * 6) + (definedNames.Count > 0 ? definedNames.Count * (4 + 27) : 0);
        foreach (var sheet in sheets)
        {
            foreach (var row in sheet.Rows)
                n += row.Cells.Length * 40 + 32;
        }
        n += styles.Fonts.Count * 64 + styles.NumFmts.Count * 32 + styles.CellXfs.Count * 24;
        return Math.Max(n, 256 * 1024);
    }

    private static int WriteBiff(Span<byte> buffer, List<ReadOnlyMemory<char>> sharedStrings, List<SheetData> sheets, StylesData styles, List<DefinedNameInfo> definedNames, Action<string>? log)
    {
        var pos = 0;
        var bw = new BiffWriter(buffer);
        bw.WriteBofWorkbook();
        bw.WriteCodepage(0x04E4);

        foreach (var font in styles.Fonts)
        {
            var twips = (ushort)(font.Height * 20);
            bw.WriteFont(font.Name, twips, font.Bold, font.Italic);
        }
        pos = bw.Position;

        bw.WriteBuiltinFmtCount(0);
        for (var i = 0; i < styles.NumFmts.Count; i++)
        {
            var nf = styles.NumFmts[i];
            bw.WriteFormat(164 + i, nf.FormatCode);
        }
        pos = bw.Position;

        for (var i = 0; i < 15; i++)
            bw.WriteXf(0, 0, false, null);
        foreach (var xf in styles.CellXfs)
        {
            var fontIdx = styles.GetBiffFontIndex(xf.FontId);
            var fmtIdx = styles.GetBiffFormatIndex(xf.NumFmtId);
            bw.WriteXf(fontIdx, fmtIdx, true, xf);
        }
        pos = bw.Position;

        if (sharedStrings.Count > 0)
        {
            var sstSize = WriteSstCorrect(buffer.Slice(pos), sharedStrings);
            pos += sstSize;
        }

        var sheetIndexByName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < sheets.Count; i++)
            if (!sheetIndexByName.ContainsKey(sheets[i].Name))
                sheetIndexByName[sheets[i].Name] = i;

        var cp1252 = Encoding.GetEncoding(1252);
        var tempBuf = ArrayPool<byte>.Shared.Rent(64 * 1024);
        var sheetSizes = new List<int>();
        try
        {
            for (var i = 0; i < sheets.Count; i++)
            {
                var sz = WriteSheet(tempBuf.AsSpan(), sheets[i], sharedStrings, styles, i, sheetIndexByName, log);
                sheetSizes.Add(sz);
            }

            var boundsheetLen = 0;
            foreach (var sheet in sheets)
            {
                var name = TruncateSheetName(sheet.Name);
                var nameBytes = Math.Min(cp1252.GetByteCount(name), 31);
                boundsheetLen += 4 + 4 + 1 + 1 + 1 + nameBytes;
            }
            var externNameLen = sheets.Count > 0 ? (8 + (4 + 2 + sheets.Count * 6) + definedNames.Count * (4 + 27)) : 0;
            var eofLen = 4;
            var firstSheetOffset = pos + boundsheetLen + externNameLen + eofLen;

            var offset = firstSheetOffset;
            for (var i = 0; i < sheets.Count; i++)
            {
                var sheet = sheets[i];
                bw = new BiffWriter(buffer.Slice(pos));
                bw.WriteBoundSheet(offset, TruncateSheetName(sheet.Name), sheet.Visibility);
                var name = TruncateSheetName(sheet.Name);
                var nameBytes = Math.Min(cp1252.GetByteCount(name), 31);
                pos += 4 + 4 + 1 + 1 + 1 + nameBytes;
                offset += sheetSizes[i];
            }

            if (sheets.Count > 0)
            {
                bw = new BiffWriter(buffer.Slice(pos));
                bw.WriteSupBookInternalRef(sheets.Count);
                pos += 8;
                bw = new BiffWriter(buffer.Slice(pos));
                bw.WriteExternSheet(sheets.Count);
                pos += 4 + 2 + sheets.Count * 6;
                if (definedNames.Count > 0)
                    foreach (var dn in definedNames)
                    {
                        bw = new BiffWriter(buffer.Slice(pos));
                        bw.WriteNameBuiltin(dn, (ushort)dn.SheetIndex0Based);
                        pos += 4 + 27;
                    }
            }

            bw = new BiffWriter(buffer.Slice(pos));
            bw.WriteEof();
            pos += 4;

            for (var i = 0; i < sheets.Count; i++)
            {
                var sz = WriteSheet(tempBuf.AsSpan(), sheets[i], sharedStrings, styles, i, sheetIndexByName, log);
                tempBuf.AsSpan(0, sz).CopyTo(buffer.Slice(pos));
                pos += sz;
            }
        }
        finally
        {
            ArrayPool<byte>.Shared.Return(tempBuf);
        }

        return pos;
    }

    private static (int FirstRow, int LastRowPlus1, int FirstCol, int LastColPlus1) ComputeUsedRange(SheetData sheet)
    {
        var firstRow = int.MaxValue;
        var lastRow = -1;
        var firstCol = int.MaxValue;
        var lastCol = -1;
        foreach (var row in sheet.Rows)
        {
            if (row.RowIndex < firstRow) firstRow = row.RowIndex;
            if (row.RowIndex > lastRow) lastRow = row.RowIndex;
            foreach (var c in row.Cells)
            {
                if (c.Col < firstCol) firstCol = c.Col;
                if (c.Col > lastCol) lastCol = c.Col;
            }
        }
        if (firstRow == int.MaxValue) firstRow = 0;
        if (lastRow < 0) lastRow = 0;
        if (firstCol == int.MaxValue) firstCol = 0;
        if (lastCol < 0) lastCol = 0;
        return (firstRow, lastRow + 1, firstCol, lastCol + 1);
    }

    private static byte MapXlsxErrorToBiff(string value)
    {
        return value switch
        {
            "#NULL!" => 0x00,
            "#DIV/0!" => 0x07,
            "#VALUE!" => 0x0F,
            "#REF!" => 0x17,
            "#NAME?" => 0x1D,
            "#NUM!" => 0x24,
            "#N/A" => 0x2A,
            _ => 0x2A
        };
    }

    private static string TruncateSheetName(string name) =>
        name.Length > 31 ? name[..31] : name;

    private const int BiffMaxRecordData = 8224;

    private static int WriteSstCorrect(Span<byte> buffer, List<ReadOnlyMemory<char>> strings)
    {
        var outPos = 0;
        var dataLen = 8;
        var isSst = true;

        BinaryPrimitives.WriteInt32LittleEndian(buffer.Slice(4), strings.Count);
        BinaryPrimitives.WriteInt32LittleEndian(buffer.Slice(8), strings.Count);

        foreach (var s in strings)
        {
            var span = s.Span;
            var needs16 = false;
            foreach (var c in span)
            {
                if (c > 255) { needs16 = true; break; }
            }
            var strBytes = 3 + (needs16 ? span.Length * 2 : span.Length);
            if (dataLen + strBytes > BiffMaxRecordData)
            {
                BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(outPos), (ushort)(isSst ? 0x00FC : 0x003C));
                BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(outPos + 2), (ushort)dataLen);
                outPos += 4 + dataLen;
                dataLen = 0;
                isSst = false;
            }
            var p = outPos + 4 + dataLen;
            BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(p), (ushort)span.Length);
            buffer[p + 2] = (byte)(needs16 ? 0 : 1);
            if (needs16)
            {
                for (var i = 0; i < span.Length; i++)
                    BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(p + 3 + i * 2), span[i]);
            }
            else
            {
                for (var i = 0; i < span.Length; i++)
                    buffer[p + 3 + i] = (byte)span[i];
            }
            dataLen += strBytes;
        }

        BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(outPos), (ushort)(isSst ? 0x00FC : 0x003C));
        BinaryPrimitives.WriteUInt16LittleEndian(buffer.Slice(outPos + 2), (ushort)dataLen);
        return outPos + 4 + dataLen;
    }

    internal static int WriteSheet(Span<byte> buffer, SheetData sheet, List<ReadOnlyMemory<char>> sharedStrings, StylesData styles, int sheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName, Action<string>? log)
    {
        var bw = new BiffWriter(buffer);
        bw.WriteBofWorksheet();
        bw.WriteWsBool();
        bw.WriteDefColWidth(8);

        foreach (var col in sheet.ColInfos)
            bw.WriteColInfo(col.FirstCol, col.LastCol, col.Width > 0 ? col.Width : 10, col.Hidden);

        var (firstRow, lastRowPlus1, firstCol, lastColPlus1) = ComputeUsedRange(sheet);
        bw.WriteDimension(firstRow, lastRowPlus1, firstCol, lastColPlus1);

        foreach (var row in sheet.Rows)
        {
            if (row.Height > 0 || row.Hidden)
                bw.WriteRow(row.RowIndex, row.Height, row.Hidden);
            foreach (var cell in row.Cells)
            {
                var xfIdx = (ushort)styles.GetBiffXfIndex(cell.StyleIndex);
                switch (cell.Kind)
                {
                    case CellKind.Formula:
                    {
                        byte[] rgce;
                        try
                        {
                            rgce = CompileFormulaTokens(cell.Formula, sheetIndex, sheetIndexByName);
                        }
                        catch (Exception ex)
                        {
                            log?.Invoke($"Formula compilation failed at sheet {sheetIndex}, cell {cell.Row + 1},{cell.Col + 1}: {ex.Message}");
                            rgce = Array.Empty<byte>();
                        }
                        if (rgce.Length == 0)
                        {
                            // Fallback: emit cached value only
                            WriteCachedValue(bw, cell, xfIdx, sharedStrings);
                            break;
                        }
                        var cachedKind = cell.CachedKind;
                        var cachedNum = 0d;
                        var cachedBool = false;
                        var cachedErr = (byte)0;
                        var cachedStr = "";
                        if (cachedKind == CellKind.Number && cell.Value != null &&
                            double.TryParse(cell.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var n))
                            cachedNum = n;
                        else if (cachedKind == CellKind.Boolean && cell.Value != null)
                            cachedBool = cell.Value == "1";
                        else if (cachedKind == CellKind.Error && cell.Value != null)
                            cachedErr = MapXlsxErrorToBiff(cell.Value);
                        else if (cachedKind == CellKind.String && cell.Value != null)
                            cachedStr = cell.Value;
                        else if (cachedKind == CellKind.SharedString && cell.SstIndex >= 0 && cell.SstIndex < sharedStrings.Count)
                            cachedStr = sharedStrings[cell.SstIndex].ToString();
                        bw.WriteFormula(cell.Row, cell.Col, xfIdx, rgce, cachedKind, cachedNum, cachedBool, cachedErr, cachedStr.AsSpan());
                        break;
                    }
                    case CellKind.Empty:
                        bw.WriteBlank(cell.Row, cell.Col, xfIdx);
                        break;
                    case CellKind.Number when cell.Value != null &&
                        double.TryParse(cell.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var num):
                        bw.WriteNumber(cell.Row, cell.Col, num, xfIdx);
                        break;
                    case CellKind.Boolean when cell.Value != null:
                        bw.WriteBool(cell.Row, cell.Col, cell.Value == "1", xfIdx);
                        break;
                    case CellKind.Error when cell.Value != null:
                        bw.WriteError(cell.Row, cell.Col, MapXlsxErrorToBiff(cell.Value), xfIdx);
                        break;
                    case CellKind.SharedString when cell.SstIndex >= 0 && cell.SstIndex < sharedStrings.Count:
                        bw.WriteLabelSst(cell.Row, cell.Col, cell.SstIndex, xfIdx);
                        break;
                    default:
                        if (cell.Value != null || (cell.SstIndex >= 0 && cell.SstIndex < sharedStrings.Count))
                        {
                            var text = cell.Value ?? sharedStrings[cell.SstIndex].ToString();
                            bw.WriteLabel(cell.Row, cell.Col, text.AsSpan(), xfIdx);
                        }
                        break;
                }
            }
        }

        if (sheet.MergeRanges.Count > 0)
        {
            var arr = sheet.MergeRanges.Select(m => (m.FirstRow, m.FirstCol, m.LastRow, m.LastCol)).ToArray();
            bw.WriteMergeCells(arr);
        }

        if (sheet.RowBreaks.Count > 0)
        {
            var breaks = sheet.RowBreaks.OrderBy(x => x).Distinct().Select(r => (r, 0, 255)).ToArray();
            bw.WriteHorizontalPageBreaks(breaks);
        }
        if (sheet.ColBreaks.Count > 0)
        {
            var breaks = sheet.ColBreaks.OrderBy(x => x).Distinct().Select(c => (c, 0, 16383)).ToArray();
            bw.WriteVerticalPageBreaks(breaks);
        }
        if (sheet.PageMargins is { } margins)
        {
            bw.WriteLeftMargin(margins.Left);
            bw.WriteRightMargin(margins.Right);
            bw.WriteTopMargin(margins.Top);
            bw.WriteBottomMargin(margins.Bottom);
        }
        if (sheet.PageSetup is { } setup)
        {
            bw.WritePageSetup(setup.Landscape, setup.Scale, setup.StartPageNumber, setup.FitToWidth, setup.FitToHeight, sheet.PageMargins?.Header ?? 0.3, sheet.PageMargins?.Footer ?? 0.3);
        }
        if (sheet.PrintOptions is { } po)
        {
            if (po.PrintGridLines)
                bw.WritePrintGridLines(true);
            if (po.PrintHeadings)
                bw.WritePrintHeaders(true);
            if (po.CenterHorizontally)
                bw.WriteCenterHorizontal(true);
            if (po.CenterVertically)
                bw.WriteCenterVertical(true);
        }
        if (sheet.HeaderFooter is { } hf)
        {
            if (!string.IsNullOrEmpty(hf.Header))
                bw.WriteHeader(hf.Header.AsSpan());
            if (!string.IsNullOrEmpty(hf.Footer))
                bw.WriteFooter(hf.Footer.AsSpan());
        }
        if (sheet.FreezePane is { } fp)
        {
            bw.WriteWindow2(freezePanes: true);
            var activePane = (byte)((fp.RowSplit > 0 && fp.ColSplit > 0) ? 0 : (fp.RowSplit > 0 ? 2 : (fp.ColSplit > 0 ? 1 : 0)));
            bw.WritePane((ushort)fp.ColSplit, (ushort)fp.RowSplit, (ushort)fp.TopRowVisible, (ushort)fp.LeftColVisible, activePane);
        }
        else
            bw.WriteWindow2(freezePanes: false);

        foreach (var hl in sheet.Hyperlinks)
            bw.WriteHyperlink(hl.FirstRow, hl.FirstCol, hl.LastRow, hl.LastCol, hl.Url.AsSpan());

        ushort shapeId = 1;
        foreach (var comment in sheet.Comments)
        {
            bw.WriteObjNote(shapeId);
            if (!string.IsNullOrEmpty(comment.Text))
                bw.WriteTxoWithText(comment.Text.AsSpan());
            bw.WriteNote(comment.Row, comment.Col, false, shapeId, comment.Author.AsSpan());
            shapeId++;
        }

        // 写入图表对象
        foreach (var chart in sheet.Charts)
        {
            // 使用ArrayPool减少内存分配
            var chartWriter = ChartWriter.CreatePooled(out var chartBuffer, 65536);
            try
            {
                var chartDataLen = chartWriter.WriteChartStream(chart, sheetIndex);

                // 写入MSODRAWING记录
                bw.WriteMsodrawingChart(shapeId, chartBuffer.AsSpan(0, chartDataLen));

                // 写入OBJ记录
                bw.WriteObjChart(shapeId, chart);

                shapeId++;
            }
            finally
            {
                chartWriter.Dispose();
            }
        }

        if (sheet.DataValidations.Count > 0)
        {
            bw.WriteDatavalidations(sheet.DataValidations.Count);
            foreach (var dv in sheet.DataValidations)
            {
                var f1 = CompileDataValidationFormula(dv, dv.Formula1, sheetIndex, sheetIndexByName);
                var f2 = CompileFormulaTokens(dv.Formula2, sheetIndex, sheetIndexByName);
                bw.WriteDatavalidation(dv, f1, f2);
            }
        }

        // Write sheet protection if enabled
        if (sheet.SheetProtection is { } protection && protection.IsProtected)
        {
            var protectionWriter = new SheetProtectionWriter(buffer.Slice(bw.Position));
            protectionWriter.WriteAllProtectionRecords(protection);
            // Update BiffWriter position
            bw = new BiffWriter(buffer.Slice(0, bw.Position + protectionWriter.Position));
        }

        // Write shapes if present
        if (sheet.Shapes.Count > 0)
        {
            var shapeWriter = ShapeWriter.CreatePooled(out var shapeBuffer, 65536);
            try
            {
                var shapeDataLen = shapeWriter.WriteAllShapes(sheet.Shapes, shapeId);
                bw.WriteMsodrawingShapes(shapeBuffer.AsSpan(0, shapeDataLen));
                shapeId += (ushort)sheet.Shapes.Count;
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(shapeBuffer);
            }
        }

        bw.WriteEof();
        return bw.Position;
    }

    private static void WriteCachedValue(BiffWriter bw, CellData cell, ushort xfIdx, List<ReadOnlyMemory<char>> sharedStrings)
    {
        switch (cell.CachedKind)
        {
            case CellKind.Number when cell.Value != null &&
                double.TryParse(cell.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var num):
                bw.WriteNumber(cell.Row, cell.Col, num, xfIdx);
                return;
            case CellKind.Boolean when cell.Value != null:
                bw.WriteBool(cell.Row, cell.Col, cell.Value == "1", xfIdx);
                return;
            case CellKind.Error when cell.Value != null:
                bw.WriteError(cell.Row, cell.Col, MapXlsxErrorToBiff(cell.Value), xfIdx);
                return;
            case CellKind.SharedString when cell.SstIndex >= 0 && cell.SstIndex < sharedStrings.Count:
                bw.WriteLabel(cell.Row, cell.Col, sharedStrings[cell.SstIndex].Span, xfIdx);
                return;
            case CellKind.String:
            default:
                if (cell.Value != null)
                    bw.WriteLabel(cell.Row, cell.Col, cell.Value.AsSpan(), xfIdx);
                else
                    bw.WriteBlank(cell.Row, cell.Col, xfIdx);
                return;
        }
    }

    internal static byte[] CompileFormulaTokens(string? formula, int sheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName)
    {
        if (string.IsNullOrWhiteSpace(formula)) return Array.Empty<byte>();
        // any exceptions are allowed to bubble so callers/consumers can react or log
        return FormulaCompiler.Compile(formula, sheetIndex, sheetIndexByName);
    }

    private static byte[] CompileDataValidationFormula(DataValidationInfo dv, string formula, int sheetIndex, IReadOnlyDictionary<string, int> sheetIndexByName)
    {
        if (string.IsNullOrWhiteSpace(formula)) return [];
        // List type: keep explicit list as ptgStr (tStr)
        if (dv.Type == 3 && formula.IndexOf(',') >= 0 && !formula.Contains('!') && !formula.Contains(':') && !formula.StartsWith("=", StringComparison.Ordinal))
            return FormulaCompiler.BuildExplicitListToken(formula);
        return CompileFormulaTokens(formula, sheetIndex, sheetIndexByName);
    }
}
