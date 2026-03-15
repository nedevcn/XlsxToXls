using System;
using System.Collections.Generic;
using System.IO;
using Nedev.FileConverters.XlsxToXls;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    /// <summary>
    /// Edge case and boundary condition tests for XLSX to XLS conversion.
    /// </summary>
    public class EdgeCaseTests
    {
        #region Null and Empty Tests

        [Fact]
        public void ChartData_EmptySeries()
        {
            var chart = new ChartData
            {
                Name = "EmptyChart",
                Type = ChartType.Column,
                Series = new List<ChartSeries>()
            };

            Assert.Empty(chart.Series);
            Assert.Equal("EmptyChart", chart.Name);
        }

        [Fact]
        public void ChartSeries_NullRanges()
        {
            var series = new ChartSeries
            {
                Name = "TestSeries",
                Categories = null,
                Values = null
            };

            Assert.Null(series.Categories);
            Assert.Null(series.Values);
            Assert.Equal("TestSeries", series.Name);
        }

        [Fact]
        public void ChartTitle_EmptyText()
        {
            var title = new ChartTitle { Text = "" };
            Assert.Equal("", title.Text);
        }

        [Fact]
        public void ChartTitle_NullText()
        {
            var title = new ChartTitle { Text = null };
            Assert.Null(title.Text);
        }

        [Fact]
        public void ChartPosition_ZeroDimensions()
        {
            var position = new ChartPosition { X = 0, Y = 0, Width = 0, Height = 0 };
            
            Assert.Equal(0, position.X);
            Assert.Equal(0, position.Y);
            Assert.Equal(0, position.Width);
            Assert.Equal(0, position.Height);
        }

        [Fact]
        public void ChartPosition_NegativeCoordinates()
        {
            var position = new ChartPosition { X = -100, Y = -50, Width = 400, Height = 300 };
            
            Assert.Equal(-100, position.X);
            Assert.Equal(-50, position.Y);
        }

        #endregion

        #region Boundary Value Tests

        [Fact]
        public void ChartRange_MaximumValues()
        {
            var range = new ChartRange
            {
                SheetName = "Sheet1",
                FirstRow = 65535,
                FirstCol = 255,
                LastRow = 65535,
                LastCol = 255
            };

            Assert.Equal(65535, range.FirstRow);
            Assert.Equal(255, range.FirstCol);
            Assert.Equal(65535, range.LastRow);
            Assert.Equal(255, range.LastCol);
        }

        [Fact]
        public void ChartRange_ZeroValues()
        {
            var range = new ChartRange
            {
                SheetName = "Sheet1",
                FirstRow = 0,
                FirstCol = 0,
                LastRow = 0,
                LastCol = 0
            };

            Assert.Equal(0, range.FirstRow);
            Assert.Equal(0, range.FirstCol);
            Assert.Equal(0, range.LastRow);
            Assert.Equal(0, range.LastCol);
        }

        [Fact]
        public void ChartRange_ReversedCoordinates()
        {
            var range = new ChartRange
            {
                SheetName = "Sheet1",
                FirstRow = 10,
                FirstCol = 5,
                LastRow = 0,
                LastCol = 0
            };

            // Should allow reversed coordinates (Last < First)
            Assert.True(range.LastRow < range.FirstRow);
            Assert.True(range.LastCol < range.FirstCol);
        }

        [Fact]
        public void ChartRange_EmptySheetName()
        {
            var range = new ChartRange
            {
                SheetName = "",
                FirstRow = 0,
                FirstCol = 0,
                LastRow = 10,
                LastCol = 5
            };

            Assert.Equal("", range.SheetName);
        }

        #endregion

        #region Data Point Tests

        [Fact]
        public void ChartDataPoint_MaximumIndex()
        {
            var point = new ChartDataPoint
            {
                Index = int.MaxValue,
                FillColor = ChartColor.Red
            };

            Assert.Equal(int.MaxValue, point.Index);
        }

        [Fact]
        public void ChartDataPoint_NegativeIndex()
        {
            var point = new ChartDataPoint
            {
                Index = -1,
                FillColor = ChartColor.Blue
            };

            Assert.Equal(-1, point.Index);
        }

        [Fact]
        public void ChartDataPoint_NullColor()
        {
            var point = new ChartDataPoint
            {
                Index = 0,
                FillColor = null
            };

            Assert.Null(point.FillColor);
        }

        [Fact]
        public void ChartSeries_EmptyDataPoints()
        {
            var series = new ChartSeries
            {
                Name = "Test",
                DataPoints = new List<ChartDataPoint>()
            };

            Assert.Empty(series.DataPoints);
        }

        [Fact]
        public void ChartSeries_LargeNumberOfDataPoints()
        {
            var series = new ChartSeries
            {
                Name = "LargeSeries",
                DataPoints = new List<ChartDataPoint>()
            };

            for (int i = 0; i < 1000; i++)
            {
                series.DataPoints.Add(new ChartDataPoint
                {
                    Index = i,
                    FillColor = ChartColor.Green
                });
            }

            Assert.Equal(1000, series.DataPoints.Count);
        }

        #endregion

        #region Chart Type Tests

        [Theory]
        [InlineData(ChartType.Column)]
        [InlineData(ChartType.Bar)]
        [InlineData(ChartType.Line)]
        [InlineData(ChartType.Pie)]
        [InlineData(ChartType.Scatter)]
        [InlineData(ChartType.Area)]
        [InlineData(ChartType.Doughnut)]
        [InlineData(ChartType.Radar)]
        public void ChartType_AllTypesSupported(ChartType chartType)
        {
            var chart = new ChartData
            {
                Name = "Test",
                Type = chartType,
                Series = new List<ChartSeries>()
            };

            Assert.Equal(chartType, chart.Type);
        }

        #endregion

        #region Color Tests

        [Fact]
        public void ChartColor_DefaultValues()
        {
            var color = new ChartColor();
            
            Assert.Equal(0, color.R);
            Assert.Equal(0, color.G);
            Assert.Equal(0, color.B);
        }

        [Fact]
        public void ChartColor_MaximumValues()
        {
            var color = new ChartColor
            {
                R = 255,
                G = 255,
                B = 255
            };

            Assert.Equal(255, color.R);
            Assert.Equal(255, color.G);
            Assert.Equal(255, color.B);
        }

        [Fact]
        public void ChartColor_PredefinedColors()
        {
            // ChartColor is a struct, so we check the values directly
            Assert.Equal(new ChartColor(255, 0, 0), ChartColor.Red);
            Assert.Equal(new ChartColor(0, 255, 0), ChartColor.Green);
            Assert.Equal(new ChartColor(0, 0, 255), ChartColor.Blue);
            Assert.Equal(new ChartColor(0, 0, 0), ChartColor.Black);
            Assert.Equal(new ChartColor(255, 255, 255), ChartColor.White);
            Assert.Equal(new ChartColor(255, 255, 0), ChartColor.Yellow);
            Assert.Equal(new ChartColor(0, 255, 255), ChartColor.Cyan);
            Assert.Equal(new ChartColor(255, 0, 255), ChartColor.Magenta);
        }

        #endregion

        #region Axis Tests

        [Fact]
        public void ChartAxis_DefaultValues()
        {
            var axis = new ChartAxis();

            Assert.Null(axis.Title);
            Assert.True(axis.HasMajorGridlines);
            Assert.False(axis.HasMinorGridlines);
            Assert.Null(axis.MinValue);
            Assert.Null(axis.MaxValue);
        }

        [Fact]
        public void ChartAxis_ExtremeValues()
        {
            var axis = new ChartAxis
            {
                Title = "Test Axis",
                HasMajorGridlines = true,
                MinValue = double.MinValue,
                MaxValue = double.MaxValue
            };

            Assert.Equal(double.MinValue, axis.MinValue);
            Assert.Equal(double.MaxValue, axis.MaxValue);
        }

        [Fact]
        public void ChartAxis_InvertedRange()
        {
            var axis = new ChartAxis
            {
                MinValue = 100,
                MaxValue = 0
            };

            Assert.True(axis.MinValue > axis.MaxValue);
        }

        #endregion

        #region Legend Tests

        [Fact]
        public void ChartLegend_DefaultValues()
        {
            var legend = new ChartLegend();

            Assert.True(legend.Show);
            Assert.Equal(LegendPosition.Right, legend.Position);
        }

        [Fact]
        public void ChartLegend_AllPositions()
        {
            var positions = new[]
            {
                LegendPosition.Bottom,
                LegendPosition.Left,
                LegendPosition.Right,
                LegendPosition.Top
            };

            foreach (var position in positions)
            {
                var legend = new ChartLegend { Position = position };
                Assert.Equal(position, legend.Position);
            }
        }

        #endregion

        #region Data Labels Tests

        [Fact]
        public void DataLabels_DefaultValues()
        {
            var labels = new DataLabels();

            Assert.True(labels.Show);
            Assert.True(labels.ShowValue);
            Assert.False(labels.ShowCategory);
            Assert.False(labels.ShowPercentage);
        }

        [Fact]
        public void DataLabels_AllEnabled()
        {
            var labels = new DataLabels
            {
                Show = true,
                ShowValue = true,
                ShowCategory = true,
                ShowPercentage = true
            };

            Assert.True(labels.Show);
            Assert.True(labels.ShowValue);
            Assert.True(labels.ShowCategory);
            Assert.True(labels.ShowPercentage);
        }

        #endregion

        #region Trend Line Tests

        [Fact]
        public void TrendLine_DefaultValues()
        {
            var trendLine = new TrendLine();

            Assert.Equal(TrendLineType.Linear, trendLine.Type);
            Assert.False(trendLine.DisplayEquation);
            Assert.False(trendLine.DisplayRSquared);
        }

        [Fact]
        public void TrendLine_AllTypes()
        {
            var types = new[]
            {
                TrendLineType.Linear,
                TrendLineType.Exponential,
                TrendLineType.Logarithmic,
                TrendLineType.Polynomial,
                TrendLineType.Power,
                TrendLineType.MovingAverage
            };

            foreach (var type in types)
            {
                var trendLine = new TrendLine { Type = type };
                Assert.Equal(type, trendLine.Type);
            }
        }

        #endregion

        #region Error Bars Tests

        [Fact]
        public void ErrorBars_DefaultValues()
        {
            var errorBars = new ErrorBars();

            Assert.Equal(ErrorBarType.Both, errorBars.Type);
            Assert.Equal(ErrorBarValueType.FixedValue, errorBars.ValueType);
            Assert.Equal(0, errorBars.Value);
            Assert.True(errorBars.ShowCap);
        }

        [Fact]
        public void ErrorBars_NegativeValue()
        {
            var errorBars = new ErrorBars
            {
                ValueType = ErrorBarValueType.FixedValue,
                Value = -5.5
            };

            Assert.Equal(-5.5, errorBars.Value);
        }

        #endregion

        #region Line Style Tests

        [Fact]
        public void LineStyle_AllValues()
        {
            var styles = new[]
            {
                LineStyle.Solid,
                LineStyle.Dash,
                LineStyle.Dot,
                LineStyle.DashDot,
                LineStyle.DashDotDot,
                LineStyle.None
            };

            foreach (var style in styles)
            {
                // LineStyle is an enum, just verify it can be assigned
                var trendLine = new TrendLine { LineStyle = style };
                Assert.Equal(style, trendLine.LineStyle);
            }
        }

        #endregion

        #region Combined Chart Tests

        [Fact]
        public void ChartSeries_SecondaryChartType()
        {
            var series = new ChartSeries
            {
                Name = "ComboSeries",
                SecondaryChartType = ChartType.Line,
                UseSecondaryAxis = true
            };

            Assert.Equal(ChartType.Line, series.SecondaryChartType);
            Assert.True(series.UseSecondaryAxis);
        }

        [Fact]
        public void ChartSeries_NullSecondaryChartType()
        {
            var series = new ChartSeries
            {
                Name = "SingleTypeSeries",
                SecondaryChartType = null,
                UseSecondaryAxis = false
            };

            Assert.Null(series.SecondaryChartType);
            Assert.False(series.UseSecondaryAxis);
        }

        #endregion

        #region Marker Style Tests

        [Fact]
        public void ChartSeries_AllMarkerStyles()
        {
            var styles = new[]
            {
                MarkerStyle.None,
                MarkerStyle.Circle,
                MarkerStyle.Diamond,
                MarkerStyle.Square,
                MarkerStyle.Triangle,
                MarkerStyle.Star,
                MarkerStyle.X,
                MarkerStyle.Plus
            };

            foreach (var style in styles)
            {
                var series = new ChartSeries { MarkerStyle = style };
                Assert.Equal(style, series.MarkerStyle);
            }
        }

        #endregion

        #region Chart Writer Tests

        [Fact]
        public void ChartWriter_ValidChart()
        {
            var buffer = new byte[4096];
            var writer = new ChartWriter(buffer.AsSpan());
            var chart = new ChartData
            {
                Name = "Test",
                Type = ChartType.Column,
                Series = new List<ChartSeries>()
            };

            var bytesWritten = writer.WriteChartStream(chart, 0);
            Assert.True(bytesWritten > 0);
        }

        #endregion
    }
}
