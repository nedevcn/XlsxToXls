using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Nedev.FileConverters.XlsxToXls;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;
using Xunit.Abstractions;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    /// <summary>
    /// Performance benchmarks for XLSX to XLS conversion.
    /// These tests measure conversion speed and memory usage.
    /// </summary>
    public class PerformanceTests
    {
        private readonly ITestOutputHelper _output;

        public PerformanceTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void ChartWriter_Performance()
        {
            const int seriesCount = 10;
            const int pointsPerSeries = 100;
            
            var chart = new ChartData
            {
                Name = "PerformanceChart",
                Type = ChartType.Line,
                Title = new ChartTitle { Text = "Performance Test Chart" },
                Series = new List<ChartSeries>()
            };

            for (int i = 0; i < seriesCount; i++)
            {
                var series = new ChartSeries
                {
                    Name = $"Series{i}",
                    SeriesIndex = i,
                    DataPoints = new List<ChartDataPoint>()
                };

                for (int j = 0; j < pointsPerSeries; j++)
                {
                    series.DataPoints.Add(new ChartDataPoint
                    {
                        Index = j,
                        FillColor = ChartColor.Blue
                    });
                }

                chart.Series.Add(series);
            }

            var stopwatch = Stopwatch.StartNew();
            var writer = ChartWriter.CreatePooled(out var buffer, 1024 * 1024);
            try
            {
                var bytesWritten = writer.WriteChartStream(chart, 0);
                stopwatch.Stop();
                
                _output.WriteLine($"Chart with {seriesCount} series x {pointsPerSeries} points written in {stopwatch.ElapsedMilliseconds}ms");
                _output.WriteLine($"Bytes written: {bytesWritten}");
                
                Assert.True(bytesWritten > 0);
                Assert.True(stopwatch.ElapsedMilliseconds < 1000, "Chart writing took too long");
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void BiffWriter_Performance()
        {
            const int recordCount = 10000;
            
            var buffer = new byte[1024 * 1024 * 10]; // 10MB buffer
            var writer = new BiffWriter(buffer);
            
            var stopwatch = Stopwatch.StartNew();
            
            for (int i = 0; i < recordCount; i++)
            {
                writer.WriteNumber(i, 0, i * 1.5);
            }
            
            stopwatch.Stop();
            
            _output.WriteLine($"Wrote {recordCount} number records in {stopwatch.ElapsedMilliseconds}ms");
            _output.WriteLine($"Average: {stopwatch.ElapsedMilliseconds / (double)recordCount:F4}ms per record");
            _output.WriteLine($"Total bytes written: {writer.Position}");
            
            Assert.True(stopwatch.ElapsedMilliseconds < 5000, "BIFF writing took too long");
        }

        [Fact]
        public void ChartDataModel_MemoryUsage()
        {
            const int chartCount = 100;
            const int seriesPerChart = 5;
            const int pointsPerSeries = 50;
            
            GC.Collect();
            GC.WaitForPendingFinalizers();
            var memoryBefore = GC.GetTotalMemory(true);
            
            var charts = new List<ChartData>();
            
            for (int c = 0; c < chartCount; c++)
            {
                var chart = new ChartData
                {
                    Name = $"Chart{c}",
                    Type = ChartType.Column,
                    Title = new ChartTitle { Text = $"Chart Title {c}" },
                    Series = new List<ChartSeries>()
                };

                for (int s = 0; s < seriesPerChart; s++)
                {
                    var series = new ChartSeries
                    {
                        Name = $"Series{s}",
                        SeriesIndex = s,
                        DataPoints = new List<ChartDataPoint>()
                    };

                    for (int p = 0; p < pointsPerSeries; p++)
                    {
                        series.DataPoints.Add(new ChartDataPoint
                        {
                            Index = p,
                            FillColor = ChartColor.Red
                        });
                    }

                    chart.Series.Add(series);
                }

                charts.Add(chart);
            }
            
            var memoryAfter = GC.GetTotalMemory(true);
            var memoryUsed = memoryAfter - memoryBefore;
            
            _output.WriteLine($"Created {chartCount} charts with {seriesPerChart} series each ({pointsPerSeries} points per series)");
            _output.WriteLine($"Memory used: {memoryUsed / 1024} KB");
            _output.WriteLine($"Memory per chart: {memoryUsed / chartCount / 1024.0:F2} KB");
            
            // Memory usage should be reasonable
            Assert.True(memoryUsed < 100 * 1024 * 1024, "Used too much memory for chart data models");
        }

        [Fact]
        public void ArrayPool_BufferReuse()
        {
            const int iterations = 1000;
            const int bufferSize = 64 * 1024;
            
            var stopwatch = Stopwatch.StartNew();
            
            for (int i = 0; i < iterations; i++)
            {
                var buffer = System.Buffers.ArrayPool<byte>.Shared.Rent(bufferSize);
                // Simulate some work
                buffer[0] = (byte)i;
                System.Buffers.ArrayPool<byte>.Shared.Return(buffer);
            }
            
            stopwatch.Stop();
            
            _output.WriteLine($"Rented and returned {iterations} buffers in {stopwatch.ElapsedMilliseconds}ms");
            _output.WriteLine($"Average: {stopwatch.ElapsedMilliseconds / (double)iterations:F4}ms per operation");
            
            Assert.True(stopwatch.ElapsedMilliseconds < 1000, "Buffer pool operations took too long");
        }

        [Fact]
        public void StylesData_Performance()
        {
            const int styleCount = 1000;
            
            var styles = new StylesData();
            
            var stopwatch = Stopwatch.StartNew();
            
            for (int i = 0; i < styleCount; i++)
            {
                styles.Fonts.Add(new FontInfo($"Font{i}", 11 + (i % 5), i % 2 == 0, i % 3 == 0, -1));
                styles.CellXfs.Add(new CellXfInfo(
                    NumFmtId: i % 50,
                    FontId: i % 20,
                    FillId: 0,
                    BorderId: 0,
                    HorizontalAlign: (byte)(i % 4),
                    VerticalAlign: 2,
                    WrapText: i % 2 == 0,
                    Indent: (byte)(i % 8),
                    Locked: true,
                    Hidden: false));
            }
            
            stopwatch.Stop();
            
            _output.WriteLine($"Created {styleCount} styles in {stopwatch.ElapsedMilliseconds}ms");
            _output.WriteLine($"Fonts: {styles.Fonts.Count}");
            _output.WriteLine($"CellXfs: {styles.CellXfs.Count}");
            
            Assert.True(stopwatch.ElapsedMilliseconds < 1000, "Style creation took too long");
        }
    }
}
