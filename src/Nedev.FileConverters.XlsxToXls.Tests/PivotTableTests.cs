using System.Collections.Generic;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    public class PivotTableTests
    {
        [Fact]
        public void PivotTableData_DefaultValues()
        {
            var pivotTable = new PivotTableData();

            Assert.NotNull(pivotTable.Id);
            Assert.Equal("PivotTable1", pivotTable.Name);
            Assert.Equal(0, pivotTable.CacheId);
            Assert.NotNull(pivotTable.Location);
            Assert.Equal(0, pivotTable.Location.Row);
            Assert.Equal(0, pivotTable.Location.Column);
            Assert.NotNull(pivotTable.RowFields);
            Assert.Empty(pivotTable.RowFields);
            Assert.NotNull(pivotTable.ColumnFields);
            Assert.Empty(pivotTable.ColumnFields);
            Assert.NotNull(pivotTable.DataFields);
            Assert.Empty(pivotTable.DataFields);
            Assert.NotNull(pivotTable.PageFields);
            Assert.Empty(pivotTable.PageFields);
            Assert.NotNull(pivotTable.HiddenFields);
            Assert.Empty(pivotTable.HiddenFields);
            Assert.True(pivotTable.ShowRowGrandTotals);
            Assert.True(pivotTable.ShowColumnGrandTotals);
            Assert.False(pivotTable.ShowError);
            Assert.Null(pivotTable.ErrorString);
            Assert.False(pivotTable.ShowEmpty);
            Assert.Null(pivotTable.EmptyString);
            Assert.True(pivotTable.AutoFormat);
            Assert.True(pivotTable.PreserveFormatting);
            Assert.True(pivotTable.UseCustomLists);
            Assert.True(pivotTable.ShowExpandCollapseButtons);
            Assert.True(pivotTable.ShowFieldHeaders);
            Assert.False(pivotTable.OutlineForm);
            Assert.True(pivotTable.CompactRowAxis);
            Assert.True(pivotTable.CompactColumnAxis);
            Assert.Equal(1, pivotTable.OutlineIndent);
            Assert.Null(pivotTable.StyleName);
            Assert.Equal(MergeLabels.None, pivotTable.MergeLabels);
            Assert.Equal(0, pivotTable.PageWrap);
            Assert.Equal(PageFilterOrder.DownThenOver, pivotTable.PageFilterOrder);
            Assert.Null(pivotTable.SourceRange);
            Assert.Null(pivotTable.TargetSheet);
        }

        [Fact]
        public void PivotTableData_WithFields()
        {
            var pivotTable = new PivotTableData
            {
                Name = "SalesPivot",
                CacheId = 1,
                Location = new CellLocation { Row = 5, Column = 2 },
                RowFields = new List<PivotField>
                {
                    new() { FieldIndex = 0, Axis = PivotAxis.Row, Name = "Region" },
                    new() { FieldIndex = 1, Axis = PivotAxis.Row, Name = "Product" }
                },
                ColumnFields = new List<PivotField>
                {
                    new() { FieldIndex = 2, Axis = PivotAxis.Column, Name = "Year" }
                },
                DataFields = new List<PivotDataField>
                {
                    new() { FieldIndex = 3, Function = AggregationFunction.Sum, Name = "Sales" }
                },
                PageFields = new List<PivotField>
                {
                    new() { FieldIndex = 4, Axis = PivotAxis.Page, Name = "Category" }
                }
            };

            Assert.Equal("SalesPivot", pivotTable.Name);
            Assert.Equal(1, pivotTable.CacheId);
            Assert.Equal(5, pivotTable.Location.Row);
            Assert.Equal(2, pivotTable.Location.Column);
            Assert.Equal(2, pivotTable.RowFields.Count);
            Assert.Single(pivotTable.ColumnFields);
            Assert.Single(pivotTable.DataFields);
            Assert.Single(pivotTable.PageFields);
        }

        [Fact]
        public void PivotField_DefaultValues()
        {
            var field = new PivotField();

            Assert.Equal(0, field.FieldIndex);
            Assert.Equal(PivotAxis.Row, field.Axis);
            Assert.Equal(SubtotalType.None, field.Subtotal);
            Assert.True(field.SubtotalTop);
            Assert.False(field.ShowAllItems);
            Assert.False(field.InsertBlankRows);
            Assert.False(field.InsertPageBreaks);
            Assert.Equal(SortOrder.Ascending, field.SortOrder);
            Assert.False(field.AutoSort);
            Assert.Null(field.AutoSortField);
            Assert.False(field.AutoShow);
            Assert.Equal(10, field.AutoShowCount);
            Assert.Equal(AutoShowType.Top, field.AutoShowType);
            Assert.Null(field.AutoShowField);
            Assert.NotNull(field.HiddenItems);
            Assert.Empty(field.HiddenItems);
            Assert.Null(field.Name);
            Assert.Null(field.NumberFormat);
            Assert.Equal(0, field.OutlineLevel);
            Assert.True(field.Compact);
        }

        [Fact]
        public void PivotDataField_DefaultValues()
        {
            var field = new PivotDataField();

            Assert.Equal(0, field.FieldIndex);
            Assert.Equal(AggregationFunction.Sum, field.Function);
            Assert.Null(field.Name);
            Assert.Null(field.NumberFormat);
            Assert.Null(field.BaseField);
            Assert.Null(field.BaseItem);
            Assert.Equal(ShowDataAs.Normal, field.ShowDataAs);
            Assert.False(field.ShowAsPercentage);
            Assert.Equal(0, field.Position);
        }

        [Fact]
        public void PivotCacheDefinition_DefaultValues()
        {
            var cache = new PivotCacheDefinition();

            Assert.Equal(0, cache.CacheId);
            Assert.Null(cache.SourceRange);
            Assert.Null(cache.SourceSheet);
            Assert.NotNull(cache.Fields);
            Assert.Empty(cache.Fields);
            Assert.Equal(0, cache.RecordCount);
            Assert.False(cache.RefreshOnLoad);
            Assert.Equal(3, cache.CreatedVersion);
            Assert.Equal(3, cache.RefreshedVersion);
            Assert.Equal(3, cache.MinRefreshableVersion);
        }

        [Fact]
        public void PivotCacheField_DefaultValues()
        {
            var field = new PivotCacheField();

            Assert.Equal("", field.Name);
            Assert.Equal(CacheFieldType.String, field.Type);
            Assert.NotNull(field.SharedItems);
            Assert.Empty(field.SharedItems);
            Assert.Null(field.NumberFormat);
            Assert.False(field.MixedTypes);
            Assert.Equal(0, field.ItemCount);
        }

        [Fact]
        public void CellLocation_FromCell()
        {
            var location = CellLocation.FromCell(10, 5);

            Assert.Equal(10, location.Row);
            Assert.Equal(5, location.Column);
        }

        // Enum tests
        [Theory]
        [InlineData(PivotAxis.Row, 0)]
        [InlineData(PivotAxis.Column, 1)]
        [InlineData(PivotAxis.Page, 2)]
        [InlineData(PivotAxis.Data, 3)]
        [InlineData(PivotAxis.Hidden, 4)]
        public void PivotAxis_HasCorrectValues(PivotAxis axis, byte expected)
        {
            Assert.Equal(expected, (byte)axis);
        }

        [Theory]
        [InlineData(AggregationFunction.Average, 0)]
        [InlineData(AggregationFunction.Count, 1)]
        [InlineData(AggregationFunction.CountNums, 2)]
        [InlineData(AggregationFunction.Max, 3)]
        [InlineData(AggregationFunction.Min, 4)]
        [InlineData(AggregationFunction.Product, 5)]
        [InlineData(AggregationFunction.StdDev, 6)]
        [InlineData(AggregationFunction.StdDevP, 7)]
        [InlineData(AggregationFunction.Sum, 8)]
        [InlineData(AggregationFunction.Var, 9)]
        [InlineData(AggregationFunction.VarP, 10)]
        public void AggregationFunction_HasCorrectValues(AggregationFunction func, byte expected)
        {
            Assert.Equal(expected, (byte)func);
        }

        [Theory]
        [InlineData(SubtotalType.None, 0)]
        [InlineData(SubtotalType.Default, 1)]
        [InlineData(SubtotalType.Sum, 2)]
        [InlineData(SubtotalType.Count, 4)]
        [InlineData(SubtotalType.Average, 8)]
        [InlineData(SubtotalType.Max, 16)]
        [InlineData(SubtotalType.Min, 32)]
        [InlineData(SubtotalType.Product, 64)]
        [InlineData(SubtotalType.CountNums, 128)]
        [InlineData(SubtotalType.StdDev, 256)]
        [InlineData(SubtotalType.StdDevP, 512)]
        [InlineData(SubtotalType.Var, 1024)]
        [InlineData(SubtotalType.VarP, 2048)]
        public void SubtotalType_HasCorrectValues(SubtotalType subtotal, ushort expected)
        {
            Assert.Equal(expected, (ushort)subtotal);
        }

        [Theory]
        [InlineData(ShowDataAs.Normal, 0)]
        [InlineData(ShowDataAs.Difference, 1)]
        [InlineData(ShowDataAs.PercentOf, 2)]
        [InlineData(ShowDataAs.PercentDiff, 3)]
        [InlineData(ShowDataAs.RunTotal, 4)]
        [InlineData(ShowDataAs.PercentOfRow, 5)]
        [InlineData(ShowDataAs.PercentOfCol, 6)]
        [InlineData(ShowDataAs.PercentOfTotal, 7)]
        [InlineData(ShowDataAs.Index, 8)]
        public void ShowDataAs_HasCorrectValues(ShowDataAs showAs, byte expected)
        {
            Assert.Equal(expected, (byte)showAs);
        }

        [Theory]
        [InlineData(SortOrder.Ascending, 0)]
        [InlineData(SortOrder.Descending, 1)]
        public void SortOrder_HasCorrectValues(SortOrder order, byte expected)
        {
            Assert.Equal(expected, (byte)order);
        }

        [Theory]
        [InlineData(AutoShowType.Top, 0)]
        [InlineData(AutoShowType.Bottom, 1)]
        public void AutoShowType_HasCorrectValues(AutoShowType type, byte expected)
        {
            Assert.Equal(expected, (byte)type);
        }

        [Theory]
        [InlineData(MergeLabels.None, 0)]
        [InlineData(MergeLabels.Row, 1)]
        [InlineData(MergeLabels.Column, 2)]
        [InlineData(MergeLabels.Both, 3)]
        public void MergeLabels_HasCorrectValues(MergeLabels merge, byte expected)
        {
            Assert.Equal(expected, (byte)merge);
        }

        [Theory]
        [InlineData(PageFilterOrder.DownThenOver, 0)]
        [InlineData(PageFilterOrder.OverThenDown, 1)]
        public void PageFilterOrder_HasCorrectValues(PageFilterOrder order, byte expected)
        {
            Assert.Equal(expected, (byte)order);
        }

        [Theory]
        [InlineData(CacheFieldType.String, 0)]
        [InlineData(CacheFieldType.Numeric, 1)]
        [InlineData(CacheFieldType.Integer, 2)]
        [InlineData(CacheFieldType.Boolean, 3)]
        [InlineData(CacheFieldType.Date, 4)]
        public void CacheFieldType_HasCorrectValues(CacheFieldType type, byte expected)
        {
            Assert.Equal(expected, (byte)type);
        }

        // PivotTableWriter tests
        [Fact]
        public void PivotTableWriter_CreatePooled()
        {
            var writer = PivotTableWriter.CreatePooled(out var buffer, 65536);
            try
            {
                Assert.True(buffer.Length >= 65536);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PivotTableWriter_WritesEmptyPivotTable()
        {
            var writer = PivotTableWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var pivotTable = new PivotTableData
                {
                    Name = "TestPivot",
                    CacheId = 1,
                    Location = new CellLocation { Row = 5, Column = 2 }
                };

                var cache = new PivotCacheDefinition
                {
                    CacheId = 1,
                    RecordCount = 100
                };

                var bytesWritten = writer.WritePivotTable(pivotTable, cache, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PivotTableWriter_WritesPivotTableWithFields()
        {
            var writer = PivotTableWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var pivotTable = new PivotTableData
                {
                    Name = "SalesPivot",
                    CacheId = 1,
                    Location = new CellLocation { Row = 5, Column = 2 },
                    RowFields = new List<PivotField>
                    {
                        new() { FieldIndex = 0, Axis = PivotAxis.Row, Name = "Region" }
                    },
                    ColumnFields = new List<PivotField>
                    {
                        new() { FieldIndex = 1, Axis = PivotAxis.Column, Name = "Year" }
                    },
                    DataFields = new List<PivotDataField>
                    {
                        new() { FieldIndex = 2, Function = AggregationFunction.Sum, Name = "Sales" }
                    }
                };

                var cache = new PivotCacheDefinition
                {
                    CacheId = 1,
                    RecordCount = 100,
                    Fields = new List<PivotCacheField>
                    {
                        new() { Name = "Region", Type = CacheFieldType.String },
                        new() { Name = "Year", Type = CacheFieldType.Numeric },
                        new() { Name = "Sales", Type = CacheFieldType.Numeric }
                    }
                };

                var bytesWritten = writer.WritePivotTable(pivotTable, cache, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PivotTableWriter_WritesPivotTableWithPageFields()
        {
            var writer = PivotTableWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var pivotTable = new PivotTableData
                {
                    Name = "SalesPivot",
                    CacheId = 1,
                    Location = new CellLocation { Row = 5, Column = 2 },
                    PageFields = new List<PivotField>
                    {
                        new() { FieldIndex = 0, Axis = PivotAxis.Page, Name = "Category" }
                    }
                };

                var cache = new PivotCacheDefinition
                {
                    CacheId = 1,
                    RecordCount = 100,
                    Fields = new List<PivotCacheField>
                    {
                        new() { Name = "Category", Type = CacheFieldType.String }
                    }
                };

                var bytesWritten = writer.WritePivotTable(pivotTable, cache, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PivotTableWriter_WritesPivotTableWithDataFields()
        {
            var writer = PivotTableWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var pivotTable = new PivotTableData
                {
                    Name = "SalesPivot",
                    CacheId = 1,
                    Location = new CellLocation { Row = 5, Column = 2 },
                    DataFields = new List<PivotDataField>
                    {
                        new() { FieldIndex = 0, Function = AggregationFunction.Sum, Name = "Sales" },
                        new() { FieldIndex = 1, Function = AggregationFunction.Count, Name = "Orders" },
                        new() { FieldIndex = 2, Function = AggregationFunction.Average, Name = "AvgPrice" }
                    }
                };

                var cache = new PivotCacheDefinition
                {
                    CacheId = 1,
                    RecordCount = 100,
                    Fields = new List<PivotCacheField>
                    {
                        new() { Name = "Sales", Type = CacheFieldType.Numeric },
                        new() { Name = "Orders", Type = CacheFieldType.Numeric },
                        new() { Name = "AvgPrice", Type = CacheFieldType.Numeric }
                    }
                };

                var bytesWritten = writer.WritePivotTable(pivotTable, cache, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PivotTableWriter_WritesPivotTableWithOptions()
        {
            var writer = PivotTableWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var pivotTable = new PivotTableData
                {
                    Name = "SalesPivot",
                    CacheId = 1,
                    Location = new CellLocation { Row = 5, Column = 2 },
                    ShowRowGrandTotals = false,
                    ShowColumnGrandTotals = false,
                    ShowError = true,
                    ErrorString = "N/A",
                    ShowEmpty = true,
                    EmptyString = "-",
                    AutoFormat = false,
                    PreserveFormatting = false,
                    MergeLabels = MergeLabels.Row,
                    PageWrap = 5,
                    PageFilterOrder = PageFilterOrder.OverThenDown
                };

                var cache = new PivotCacheDefinition
                {
                    CacheId = 1,
                    RecordCount = 100
                };

                var bytesWritten = writer.WritePivotTable(pivotTable, cache, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PivotTableWriter_WritesPivotCache()
        {
            var writer = PivotTableWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var cache = new PivotCacheDefinition
                {
                    CacheId = 1,
                    RecordCount = 1000,
                    SourceSheet = "Data",
                    SourceRange = new CellRange { FirstRow = 0, FirstCol = 0, LastRow = 999, LastCol = 10 },
                    RefreshOnLoad = true,
                    Fields = new List<PivotCacheField>
                    {
                        new()
                        {
                            Name = "Region",
                            Type = CacheFieldType.String,
                            SharedItems = new List<string> { "North", "South", "East", "West" }
                        },
                        new()
                        {
                            Name = "Sales",
                            Type = CacheFieldType.Numeric
                        },
                        new()
                        {
                            Name = "Date",
                            Type = CacheFieldType.Date
                        }
                    }
                };

                var bytesWritten = writer.WritePivotCache(cache);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PivotTableWriter_WritesCompletePivotTable()
        {
            var writer = PivotTableWriter.CreatePooled(out var buffer, 131072);
            try
            {
                var pivotTable = new PivotTableData
                {
                    Name = "SalesReport",
                    CacheId = 1,
                    Location = new CellLocation { Row = 10, Column = 5 },
                    SourceRange = new CellRange { FirstRow = 0, FirstCol = 0, LastRow = 999, LastCol = 10 },
                    RowFields = new List<PivotField>
                    {
                        new() { FieldIndex = 0, Axis = PivotAxis.Row, Name = "Region", Subtotal = SubtotalType.None },
                        new() { FieldIndex = 1, Axis = PivotAxis.Row, Name = "Product", Subtotal = SubtotalType.None }
                    },
                    ColumnFields = new List<PivotField>
                    {
                        new() { FieldIndex = 2, Axis = PivotAxis.Column, Name = "Year", Subtotal = SubtotalType.None }
                    },
                    DataFields = new List<PivotDataField>
                    {
                        new() { FieldIndex = 3, Function = AggregationFunction.Sum, Name = "Total Sales" },
                        new() { FieldIndex = 4, Function = AggregationFunction.Count, Name = "Order Count" }
                    },
                    PageFields = new List<PivotField>
                    {
                        new() { FieldIndex = 5, Axis = PivotAxis.Page, Name = "Category" }
                    },
                    ShowRowGrandTotals = true,
                    ShowColumnGrandTotals = true,
                    ShowFieldHeaders = true
                };

                var cache = new PivotCacheDefinition
                {
                    CacheId = 1,
                    RecordCount = 1000,
                    SourceSheet = "SalesData",
                    SourceRange = new CellRange { FirstRow = 0, FirstCol = 0, LastRow = 999, LastCol = 10 },
                    Fields = new List<PivotCacheField>
                    {
                        new() { Name = "Region", Type = CacheFieldType.String, SharedItems = new List<string> { "North", "South", "East", "West" } },
                        new() { Name = "Product", Type = CacheFieldType.String, SharedItems = new List<string> { "A", "B", "C" } },
                        new() { Name = "Year", Type = CacheFieldType.Numeric, SharedItems = new List<string> { "2022", "2023", "2024" } },
                        new() { Name = "Sales", Type = CacheFieldType.Numeric },
                        new() { Name = "Orders", Type = CacheFieldType.Numeric },
                        new() { Name = "Category", Type = CacheFieldType.String, SharedItems = new List<string> { "Electronics", "Clothing", "Food" } }
                    }
                };

                // Write cache first
                var cacheBytes = writer.WritePivotCache(cache);
                Assert.True(cacheBytes > 0);

                // Then write pivot table
                var tableBytes = writer.WritePivotTable(pivotTable, cache, 0);
                Assert.True(tableBytes > 0);

                // Total should be sum of both
                Assert.True(cacheBytes + tableBytes > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }
    }
}
