using System.IO.Compression;
using System.Xml;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Reads pivot table definitions from XLSX files.
/// Parses pivotCacheDefinition and pivotTableDefinition XML files.
/// </summary>
public static class PivotTableReader
{
    private static readonly XmlReaderSettings Settings = new()
    {
        IgnoreWhitespace = true,
        IgnoreComments = true
    };

    /// <summary>
    /// Reads all pivot cache definitions from the workbook.
    /// </summary>
    public static List<PivotCacheDefinition> ReadPivotCaches(ZipArchive archive, Action<string>? log = null)
    {
        var caches = new List<PivotCacheDefinition>();

        try
        {
            // First, read the workbook.xml to find pivot cache references
            var workbookEntry = archive.GetEntry("xl/workbook.xml");
            if (workbookEntry == null)
            {
                log?.Invoke("[PivotTableReader] Workbook.xml not found");
                return caches;
            }

            log?.Invoke("[PivotTableReader] Reading pivot cache definitions");

            // Find all pivot cache definition files
            var cacheEntries = archive.Entries
                .Where(e => e.FullName.StartsWith("xl/pivotCache/pivotCacheDefinition", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var entry in cacheEntries)
            {
                var cache = ReadPivotCacheDefinition(entry, log);
                if (cache != null)
                {
                    caches.Add(cache);
                    log?.Invoke($"[PivotTableReader] Found pivot cache: {cache.CacheId}");
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PivotTableReader] Error reading pivot caches: {ex.Message}");
        }

        return caches;
    }

    /// <summary>
    /// Reads all pivot tables from the workbook.
    /// </summary>
    public static List<PivotTableData> ReadPivotTables(ZipArchive archive, Action<string>? log = null)
    {
        var pivotTables = new List<PivotTableData>();

        try
        {
            log?.Invoke("[PivotTableReader] Reading pivot tables");

            // Find all pivot table definition files
            var tableEntries = archive.Entries
                .Where(e => e.FullName.StartsWith("xl/pivotTables/pivotTableDefinition", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var entry in tableEntries)
            {
                var pivotTable = ReadPivotTableDefinition(entry, log);
                if (pivotTable != null)
                {
                    pivotTables.Add(pivotTable);
                    log?.Invoke($"[PivotTableReader] Found pivot table: {pivotTable.Name}");
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PivotTableReader] Error reading pivot tables: {ex.Message}");
        }

        return pivotTables;
    }

    private static PivotCacheDefinition? ReadPivotCacheDefinition(ZipArchiveEntry entry, Action<string>? log)
    {
        try
        {
            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, Settings);
            var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            var cache = new PivotCacheDefinition();

            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;

                if (reader.LocalName == "pivotCacheDefinition" && reader.NamespaceURI == ns)
                {
                    // Read cache ID
                    if (reader.GetAttribute("cacheId") is { } cacheIdStr &&
                        int.TryParse(cacheIdStr, out var cacheId))
                    {
                        cache.CacheId = cacheId;
                    }

                    // Read refresh on load
                    if (reader.GetAttribute("refreshOnLoad") is { } refreshStr &&
                        bool.TryParse(refreshStr, out var refresh))
                    {
                        cache.RefreshOnLoad = refresh;
                    }

                    // Read version info
                    if (reader.GetAttribute("createdVersion") is { } createdVerStr &&
                        int.TryParse(createdVerStr, out var createdVer))
                    {
                        cache.CreatedVersion = createdVer;
                    }

                    if (reader.GetAttribute("refreshedVersion") is { } refreshedVerStr &&
                        int.TryParse(refreshedVerStr, out var refreshedVer))
                    {
                        cache.RefreshedVersion = refreshedVer;
                    }

                    if (reader.GetAttribute("minRefreshableVersion") is { } minVerStr &&
                        int.TryParse(minVerStr, out var minVer))
                    {
                        cache.MinRefreshableVersion = minVer;
                    }
                }
                else if (reader.LocalName == "cacheSource" && reader.NamespaceURI == ns)
                {
                    // Read source type and range
                    if (reader.GetAttribute("type") is { } sourceType)
                    {
                        if (sourceType == "worksheet" && !reader.IsEmptyElement)
                        {
                            ReadCacheSource(reader, cache, ns, log);
                        }
                    }
                }
                else if (reader.LocalName == "cacheField" && reader.NamespaceURI == ns)
                {
                    var field = ReadCacheField(reader, ns, log);
                    if (field != null)
                    {
                        cache.Fields.Add(field);
                    }
                }
            }

            return cache;
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PivotTableReader] Error reading cache definition: {ex.Message}");
            return null;
        }
    }

    private static void ReadCacheSource(XmlReader reader, PivotCacheDefinition cache, string ns, Action<string>? log)
    {
        try
        {
            var depth = 1;
            while (reader.Read() && depth > 0)
            {
                if (reader.NodeType == XmlNodeType.Element &&
                    reader.LocalName == "worksheetSource" &&
                    reader.NamespaceURI == ns)
                {
                    // Read source sheet
                    if (reader.GetAttribute("sheet") is { } sheet)
                    {
                        cache.SourceSheet = sheet;
                    }

                    // Read source range
                    if (reader.GetAttribute("ref") is { } refRange)
                    {
                        cache.SourceRange = ParseRange(refRange);
                    }

                    // Read named range
                    if (reader.GetAttribute("name") is { } name)
                    {
                        // Named range reference
                        log?.Invoke($"[PivotTableReader] Cache uses named range: {name}");
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement &&
                         reader.LocalName == "cacheSource" &&
                         reader.NamespaceURI == ns)
                {
                    depth--;
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PivotTableReader] Error reading cache source: {ex.Message}");
        }
    }

    private static PivotCacheField? ReadCacheField(XmlReader reader, string ns, Action<string>? log)
    {
        try
        {
            var field = new PivotCacheField();

            // Read field name
            if (reader.GetAttribute("name") is { } name)
            {
                field.Name = name;
            }

            // Read number format
            if (reader.GetAttribute("numFmtId") is { } numFmtStr &&
                int.TryParse(numFmtStr, out var numFmt))
            {
                field.NumberFormat = numFmt;
            }

            // Read shared items if present
            if (!reader.IsEmptyElement)
            {
                var depth = 1;
                while (reader.Read() && depth > 0)
                {
                    if (reader.NodeType == XmlNodeType.Element &&
                        reader.NamespaceURI == ns)
                    {
                        if (reader.LocalName == "sharedItems")
                        {
                            ReadSharedItems(reader, field, ns, log);
                        }
                        else if (reader.LocalName == "fieldGroup")
                        {
                            // Field grouping - skip for now
                            if (!reader.IsEmptyElement)
                            {
                                var groupDepth = 1;
                                while (reader.Read() && groupDepth > 0)
                                {
                                    if (reader.NodeType == XmlNodeType.EndElement &&
                                        reader.LocalName == "fieldGroup" &&
                                        reader.NamespaceURI == ns)
                                    {
                                        groupDepth--;
                                    }
                                }
                            }
                        }
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement &&
                             reader.LocalName == "cacheField" &&
                             reader.NamespaceURI == ns)
                    {
                        depth--;
                    }
                }
            }

            return field;
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PivotTableReader] Error reading cache field: {ex.Message}");
            return null;
        }
    }

    private static void ReadSharedItems(XmlReader reader, PivotCacheField field, string ns, Action<string>? log)
    {
        try
        {
            // Determine field type from shared items
            if (reader.GetAttribute("containsSemiMixedTypes") is { } semiMixedStr &&
                bool.TryParse(semiMixedStr, out var semiMixed) && semiMixed)
            {
                field.MixedTypes = true;
            }

            if (reader.GetAttribute("containsString") is { } hasStringStr &&
                bool.TryParse(hasStringStr, out var hasString) && hasString)
            {
                field.Type = CacheFieldType.String;
            }

            if (reader.GetAttribute("containsNumber") is { } hasNumStr &&
                bool.TryParse(hasNumStr, out var hasNum) && hasNum)
            {
                field.Type = CacheFieldType.Numeric;
            }

            if (reader.GetAttribute("containsDate") is { } hasDateStr &&
                bool.TryParse(hasDateStr, out var hasDate) && hasDate)
            {
                field.Type = CacheFieldType.Date;
            }

            // Read shared item values
            if (!reader.IsEmptyElement)
            {
                var depth = 1;
                while (reader.Read() && depth > 0)
                {
                    if (reader.NodeType == XmlNodeType.Element &&
                        reader.NamespaceURI == ns)
                    {
                        if (reader.LocalName == "s" || reader.LocalName == "n" ||
                            reader.LocalName == "d" || reader.LocalName == "b")
                        {
                            var value = reader.GetAttribute("v") ?? "";
                            if (!string.IsNullOrEmpty(value))
                            {
                                field.SharedItems.Add(value);
                            }

                            if (!reader.IsEmptyElement)
                            {
                                reader.Read(); // Skip to end element
                            }
                        }
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement &&
                             reader.LocalName == "sharedItems" &&
                             reader.NamespaceURI == ns)
                    {
                        depth--;
                    }
                }
            }

            log?.Invoke($"[PivotTableReader] Field '{field.Name}' has {field.SharedItems.Count} shared items");
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PivotTableReader] Error reading shared items: {ex.Message}");
        }
    }

    private static PivotTableData? ReadPivotTableDefinition(ZipArchiveEntry entry, Action<string>? log)
    {
        try
        {
            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, Settings);
            var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            var pivotTable = new PivotTableData();

            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element) continue;

                if (reader.LocalName == "pivotTableDefinition" && reader.NamespaceURI == ns)
                {
                    // Read name
                    if (reader.GetAttribute("name") is { } name)
                    {
                        pivotTable.Name = name;
                    }

                    // Read cache ID
                    if (reader.GetAttribute("cacheId") is { } cacheIdStr &&
                        int.TryParse(cacheIdStr, out var cacheId))
                    {
                        pivotTable.CacheId = cacheId;
                    }

                    // Read various flags
                    if (reader.GetAttribute("showRowGrandTotals") is { } rowGrandStr &&
                        bool.TryParse(rowGrandStr, out var rowGrand))
                    {
                        pivotTable.ShowRowGrandTotals = rowGrand;
                    }

                    if (reader.GetAttribute("showColGrandTotals") is { } colGrandStr &&
                        bool.TryParse(colGrandStr, out var colGrand))
                    {
                        pivotTable.ShowColumnGrandTotals = colGrand;
                    }

                    if (reader.GetAttribute("showError") is { } showErrorStr &&
                        bool.TryParse(showErrorStr, out var showError))
                    {
                        pivotTable.ShowError = showError;
                    }

                    if (reader.GetAttribute("errorString") is { } errorStr)
                    {
                        pivotTable.ErrorString = errorStr;
                    }

                    if (reader.GetAttribute("showEmpty") is { } showEmptyStr &&
                        bool.TryParse(showEmptyStr, out var showEmpty))
                    {
                        pivotTable.ShowEmpty = showEmpty;
                    }

                    if (reader.GetAttribute("emptyString") is { } emptyStr)
                    {
                        pivotTable.EmptyString = emptyStr;
                    }

                    if (reader.GetAttribute("autoFormatId") is { } autoFmtStr &&
                        int.TryParse(autoFmtStr, out var autoFmt))
                    {
                        pivotTable.AutoFormat = autoFmt > 0;
                    }

                    if (reader.GetAttribute("applyNumberFormats") is { } numFmtStr &&
                        bool.TryParse(numFmtStr, out var numFmt))
                    {
                        pivotTable.PreserveFormatting = numFmt;
                    }

                    if (reader.GetAttribute("useAutoFormatting") is { } autoFmt2Str &&
                        bool.TryParse(autoFmt2Str, out var autoFmt2))
                    {
                        pivotTable.AutoFormat = autoFmt2;
                    }

                    if (reader.GetAttribute("indent") is { } indentStr &&
                        byte.TryParse(indentStr, out var indent))
                    {
                        pivotTable.OutlineIndent = indent;
                    }

                    if (reader.GetAttribute("outline") is { } outlineStr &&
                        bool.TryParse(outlineStr, out var outline))
                    {
                        pivotTable.OutlineForm = outline;
                    }

                    if (reader.GetAttribute("outlineData") is { } outlineDataStr &&
                        bool.TryParse(outlineDataStr, out var outlineData))
                    {
                        pivotTable.OutlineForm = outlineData;
                    }

                    if (reader.GetAttribute("compact") is { } compactStr &&
                        bool.TryParse(compactStr, out var compact))
                    {
                        pivotTable.CompactRowAxis = compact;
                        pivotTable.CompactColumnAxis = compact;
                    }

                    if (reader.GetAttribute("compactData") is { } compactDataStr &&
                        bool.TryParse(compactDataStr, out var compactData))
                    {
                        pivotTable.CompactRowAxis = compactData;
                        pivotTable.CompactColumnAxis = compactData;
                    }

                    if (reader.GetAttribute("gridDropZones") is { } gridDropStr &&
                        bool.TryParse(gridDropStr, out var gridDrop))
                    {
                        pivotTable.ShowExpandCollapseButtons = gridDrop;
                    }

                    if (reader.GetAttribute("showHeaders") is { } showHdrStr &&
                        bool.TryParse(showHdrStr, out var showHdr))
                    {
                        pivotTable.ShowFieldHeaders = showHdr;
                    }

                    if (reader.GetAttribute("mergeItem") is { } mergeStr &&
                        int.TryParse(mergeStr, out var merge))
                    {
                        pivotTable.MergeLabels = (MergeLabels)merge;
                    }

                    if (reader.GetAttribute("pageWrap") is { } pageWrapStr &&
                        int.TryParse(pageWrapStr, out var pageWrap))
                    {
                        pivotTable.PageWrap = pageWrap;
                    }

                    if (reader.GetAttribute("pageOverThenDown") is { } pageOrderStr &&
                        bool.TryParse(pageOrderStr, out var pageOrder))
                    {
                        pivotTable.PageFilterOrder = pageOrder ? PageFilterOrder.OverThenDown : PageFilterOrder.DownThenOver;
                    }

                    if (reader.GetAttribute("dataCaption") is { } dataCaption)
                    {
                        // Data field caption
                    }
                }
                else if (reader.LocalName == "location" && reader.NamespaceURI == ns)
                {
                    // Read location
                    if (reader.GetAttribute("ref") is { } locationRef)
                    {
                        var location = ParseCellReference(locationRef);
                        if (location != null)
                        {
                            pivotTable.Location = location;
                        }
                    }

                    // Read first header row
                    if (reader.GetAttribute("firstHeaderRow") is { } firstHdrStr &&
                        int.TryParse(firstHdrStr, out var firstHdr))
                    {
                        // First header row info
                    }

                    // Read first data row
                    if (reader.GetAttribute("firstDataRow") is { } firstDataStr &&
                        int.TryParse(firstDataStr, out var firstData))
                    {
                        // First data row info
                    }

                    // Read first data column
                    if (reader.GetAttribute("firstDataCol") is { } firstColStr &&
                        int.TryParse(firstColStr, out var firstCol))
                    {
                        // First data column info
                    }
                }
                else if (reader.LocalName == "pivotFields" && reader.NamespaceURI == ns)
                {
                    // Pivot fields - read field definitions
                    ReadPivotFields(reader, pivotTable, ns, log);
                }
                else if (reader.LocalName == "rowFields" && reader.NamespaceURI == ns)
                {
                    pivotTable.RowFields = ReadAxisFields(reader, ns, PivotAxis.Row, log);
                }
                else if (reader.LocalName == "colFields" && reader.NamespaceURI == ns)
                {
                    pivotTable.ColumnFields = ReadAxisFields(reader, ns, PivotAxis.Column, log);
                }
                else if (reader.LocalName == "pageFields" && reader.NamespaceURI == ns)
                {
                    pivotTable.PageFields = ReadAxisFields(reader, ns, PivotAxis.Page, log);
                }
                else if (reader.LocalName == "dataFields" && reader.NamespaceURI == ns)
                {
                    pivotTable.DataFields = ReadDataFields(reader, ns, log);
                }
            }

            return pivotTable;
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PivotTableReader] Error reading pivot table definition: {ex.Message}");
            return null;
        }
    }

    private static void ReadPivotFields(XmlReader reader, PivotTableData pivotTable, string ns, Action<string>? log)
    {
        try
        {
            if (reader.IsEmptyElement) return;

            var depth = 1;
            while (reader.Read() && depth > 0)
            {
                if (reader.NodeType == XmlNodeType.Element &&
                    reader.LocalName == "pivotField" &&
                    reader.NamespaceURI == ns)
                {
                    var field = new PivotField();

                    // Read axis
                    if (reader.GetAttribute("axis") is { } axisStr)
                    {
                        field.Axis = axisStr.ToLowerInvariant() switch
                        {
                            "axisrow" => PivotAxis.Row,
                            "axiscol" => PivotAxis.Column,
                            "axispage" => PivotAxis.Page,
                            "axisdata" => PivotAxis.Data,
                            _ => PivotAxis.Hidden
                        };
                    }

                    // Read subtotal type
                    if (reader.GetAttribute("subtotalCaption") is { } subtotalStr)
                    {
                        // Parse subtotal type
                    }

                    // Read show all
                    if (reader.GetAttribute("showAll") is { } showAllStr &&
                        bool.TryParse(showAllStr, out var showAll))
                    {
                        field.ShowAllItems = showAll;
                    }

                    // Read sort order
                    if (reader.GetAttribute("sortType") is { } sortTypeStr)
                    {
                        field.SortOrder = sortTypeStr.ToLowerInvariant() == "descending"
                            ? SortOrder.Descending
                            : SortOrder.Ascending;
                    }

                    // Read auto sort
                    if (reader.GetAttribute("autoSort") is { } autoSortStr &&
                        bool.TryParse(autoSortStr, out var autoSort))
                    {
                        field.AutoSort = autoSort;
                    }

                    // Read auto show
                    if (reader.GetAttribute("autoShow") is { } autoShowStr &&
                        bool.TryParse(autoShowStr, out var autoShow))
                    {
                        field.AutoShow = autoShow;
                    }

                    // Read hidden items
                    if (!reader.IsEmptyElement)
                    {
                        var fieldDepth = 1;
                        while (reader.Read() && fieldDepth > 0)
                        {
                            if (reader.NodeType == XmlNodeType.Element &&
                                reader.NamespaceURI == ns)
                            {
                                if (reader.LocalName == "item")
                                {
                                    if (reader.GetAttribute("h") is { } hiddenStr &&
                                        hiddenStr == "1")
                                    {
                                        if (reader.GetAttribute("x") is { } idxStr &&
                                            int.TryParse(idxStr, out var idx))
                                        {
                                            field.HiddenItems.Add(idx);
                                        }
                                    }
                                }
                            }
                            else if (reader.NodeType == XmlNodeType.EndElement &&
                                     reader.LocalName == "pivotField" &&
                                     reader.NamespaceURI == ns)
                            {
                                fieldDepth--;
                            }
                        }
                    }

                    // Add to appropriate list based on axis
                    switch (field.Axis)
                    {
                        case PivotAxis.Row:
                            // Will be added via rowFields
                            break;
                        case PivotAxis.Column:
                            // Will be added via colFields
                            break;
                        case PivotAxis.Page:
                            // Will be added via pageFields
                            break;
                        case PivotAxis.Data:
                            // Will be added via dataFields
                            break;
                        case PivotAxis.Hidden:
                            pivotTable.HiddenFields.Add(field);
                            break;
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement &&
                         reader.LocalName == "pivotFields" &&
                         reader.NamespaceURI == ns)
                {
                    depth--;
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PivotTableReader] Error reading pivot fields: {ex.Message}");
        }
    }

    private static List<PivotField> ReadAxisFields(XmlReader reader, string ns, PivotAxis axis, Action<string>? log)
    {
        var fields = new List<PivotField>();

        try
        {
            if (reader.IsEmptyElement) return fields;

            var depth = 1;
            while (reader.Read() && depth > 0)
            {
                if (reader.NodeType == XmlNodeType.Element &&
                    reader.LocalName == "field" &&
                    reader.NamespaceURI == ns)
                {
                    if (reader.GetAttribute("x") is { } idxStr &&
                        int.TryParse(idxStr, out var idx))
                    {
                        fields.Add(new PivotField
                        {
                            FieldIndex = idx,
                            Axis = axis
                        });
                    }
                }
                else if (reader.NodeType == XmlNodeType.EndElement &&
                         reader.LocalName is "rowFields" or "colFields" or "pageFields" &&
                         reader.NamespaceURI == ns)
                {
                    depth--;
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PivotTableReader] Error reading axis fields: {ex.Message}");
        }

        return fields;
    }

    private static List<PivotDataField> ReadDataFields(XmlReader reader, string ns, Action<string>? log)
    {
        var fields = new List<PivotDataField>();

        try
        {
            if (reader.IsEmptyElement) return fields;

            var depth = 1;
            while (reader.Read() && depth > 0)
            {
                if (reader.NodeType == XmlNodeType.Element &&
                    reader.LocalName == "dataField" &&
                    reader.NamespaceURI == ns)
                {
                    var field = new PivotDataField();

                    // Read field index
                    if (reader.GetAttribute("fld") is { } fldStr &&
                        int.TryParse(fldStr, out var fld))
                    {
                        field.FieldIndex = fld;
                    }

                    // Read subtotal function
                    if (reader.GetAttribute("subtotal") is { } subtotalStr)
                    {
                        field.Function = subtotalStr.ToLowerInvariant() switch
                        {
                            "average" => AggregationFunction.Average,
                            "count" => AggregationFunction.Count,
                            "countNums" => AggregationFunction.CountNums,
                            "max" => AggregationFunction.Max,
                            "min" => AggregationFunction.Min,
                            "product" => AggregationFunction.Product,
                            "stdDev" => AggregationFunction.StdDev,
                            "stdDevP" => AggregationFunction.StdDevP,
                            "sum" => AggregationFunction.Sum,
                            "var" => AggregationFunction.Var,
                            "varP" => AggregationFunction.VarP,
                            _ => AggregationFunction.Sum
                        };
                    }

                    // Read name
                    if (reader.GetAttribute("name") is { } name)
                    {
                        field.Name = name;
                    }

                    // Read show data as
                    if (reader.GetAttribute("showDataAs") is { } showAsStr)
                    {
                        field.ShowDataAs = showAsStr.ToLowerInvariant() switch
                        {
                            "normal" => ShowDataAs.Normal,
                            "difference" => ShowDataAs.Difference,
                            "percent" => ShowDataAs.PercentOf,
                            "percentDiff" => ShowDataAs.PercentDiff,
                            "runTotal" => ShowDataAs.RunTotal,
                            "percentOfRow" => ShowDataAs.PercentOfRow,
                            "percentOfCol" => ShowDataAs.PercentOfCol,
                            "percentOfTotal" => ShowDataAs.PercentOfTotal,
                            "index" => ShowDataAs.Index,
                            _ => ShowDataAs.Normal
                        };
                    }

                    // Read base field
                    if (reader.GetAttribute("baseField") is { } baseFldStr &&
                        int.TryParse(baseFldStr, out var baseFld))
                    {
                        field.BaseField = baseFld;
                    }

                    // Read base item
                    if (reader.GetAttribute("baseItem") is { } baseItemStr &&
                        int.TryParse(baseItemStr, out var baseItem))
                    {
                        field.BaseItem = baseItem;
                    }

                    // Read position
                    if (reader.GetAttribute("numFmtId") is { } numFmtStr &&
                        int.TryParse(numFmtStr, out var numFmt))
                    {
                        field.NumberFormat = numFmt;
                    }

                    fields.Add(field);
                }
                else if (reader.NodeType == XmlNodeType.EndElement &&
                         reader.LocalName == "dataFields" &&
                         reader.NamespaceURI == ns)
                {
                    depth--;
                }
            }
        }
        catch (Exception ex)
        {
            log?.Invoke($"[PivotTableReader] Error reading data fields: {ex.Message}");
        }

        return fields;
    }

    private static CellRange? ParseRange(string rangeStr)
    {
        try
        {
            rangeStr = rangeStr.Replace("$", "");

            if (rangeStr.Contains(':'))
            {
                var parts = rangeStr.Split(':', 2);
                if (parts.Length == 2)
                {
                    var (col1, row1) = ParseCellRef(parts[0]);
                    var (col2, row2) = ParseCellRef(parts[1]);

                    if (col1.HasValue && row1.HasValue && col2.HasValue && row2.HasValue)
                    {
                        return new CellRange
                        {
                            FirstRow = row1.Value,
                            FirstCol = col1.Value,
                            LastRow = row2.Value,
                            LastCol = col2.Value
                        };
                    }
                }
            }
            else
            {
                var (col, row) = ParseCellRef(rangeStr);
                if (col.HasValue && row.HasValue)
                {
                    return new CellRange
                    {
                        FirstRow = row.Value,
                        FirstCol = col.Value,
                        LastRow = row.Value,
                        LastCol = col.Value
                    };
                }
            }
        }
        catch
        {
            // Ignore parsing errors
        }

        return null;
    }

    private static CellLocation? ParseCellReference(string cellRef)
    {
        var (col, row) = ParseCellRef(cellRef.Replace("$", ""));
        if (col.HasValue && row.HasValue)
        {
            return CellLocation.FromCell(row.Value, col.Value);
        }
        return null;
    }

    private static (int? col, int? row) ParseCellRef(string cellRef)
    {
        var colStr = "";
        var rowStr = "";

        foreach (var c in cellRef)
        {
            if (char.IsLetter(c))
            {
                colStr += c;
            }
            else if (char.IsDigit(c))
            {
                rowStr += c;
            }
        }

        int? col = null;
        int? row = null;

        if (!string.IsNullOrEmpty(colStr))
        {
            col = 0;
            foreach (var c in colStr.ToUpperInvariant())
            {
                col = col * 26 + (c - 'A' + 1);
            }
            col--;
        }

        if (!string.IsNullOrEmpty(rowStr) && int.TryParse(rowStr, out var rowNum))
        {
            row = rowNum - 1;
        }

        return (col, row);
    }
}
