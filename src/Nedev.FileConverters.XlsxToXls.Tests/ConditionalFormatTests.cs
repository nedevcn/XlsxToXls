using System.Collections.Generic;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    public class ConditionalFormatTests
    {
        [Fact]
        public void ConditionalFormatData_DefaultValues()
        {
            var format = new ConditionalFormatData();

            Assert.NotNull(format.Id);
            Assert.NotNull(format.Ranges);
            Assert.Empty(format.Ranges);
            Assert.Equal(ConditionalFormatType.CellIs, format.Type);
            Assert.Equal(ComparisonOperator.GreaterThan, format.Operator);
            Assert.Equal(1, format.Priority);
            Assert.False(format.StopIfTrue);
        }

        [Fact]
        public void ConditionalFormatData_WithRanges()
        {
            var format = new ConditionalFormatData
            {
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 10, LastCol = 5 },
                    new() { FirstRow = 20, FirstCol = 0, LastRow = 30, LastCol = 5 }
                }
            };

            Assert.Equal(2, format.Ranges.Count);
            Assert.Equal(0, format.Ranges[0].FirstRow);
            Assert.Equal(20, format.Ranges[1].FirstRow);
        }

        [Fact]
        public void ConditionalFormatData_CellIsRule()
        {
            var format = new ConditionalFormatData
            {
                Type = ConditionalFormatType.CellIs,
                Operator = ComparisonOperator.GreaterThan,
                Formula1 = "100",
                Style = new ConditionalFormatStyle
                {
                    FillColor = ChartColor.Red,
                    FontColor = ChartColor.White,
                    Bold = true
                }
            };

            Assert.Equal(ConditionalFormatType.CellIs, format.Type);
            Assert.Equal(ComparisonOperator.GreaterThan, format.Operator);
            Assert.Equal("100", format.Formula1);
            Assert.Equal(ChartColor.Red, format.Style.FillColor);
            Assert.True(format.Style.Bold);
        }

        [Fact]
        public void ConditionalFormatData_BetweenRule()
        {
            var format = new ConditionalFormatData
            {
                Type = ConditionalFormatType.CellIs,
                Operator = ComparisonOperator.Between,
                Formula1 = "0",
                Formula2 = "100",
                Style = new ConditionalFormatStyle
                {
                    FillColor = ChartColor.Green
                }
            };

            Assert.Equal(ComparisonOperator.Between, format.Operator);
            Assert.Equal("0", format.Formula1);
            Assert.Equal("100", format.Formula2);
        }

        [Fact]
        public void ConditionalFormatData_ContainsTextRule()
        {
            var format = new ConditionalFormatData
            {
                Type = ConditionalFormatType.ContainsText,
                Text = "Important",
                Style = new ConditionalFormatStyle
                {
                    FillColor = ChartColor.Yellow
                }
            };

            Assert.Equal(ConditionalFormatType.ContainsText, format.Type);
            Assert.Equal("Important", format.Text);
        }

        [Fact]
        public void ConditionalFormatData_ExpressionRule()
        {
            var format = new ConditionalFormatData
            {
                Type = ConditionalFormatType.Expression,
                Formula1 = "=A1>B1",
                Style = new ConditionalFormatStyle
                {
                    FillColor = ChartColor.Blue
                }
            };

            Assert.Equal(ConditionalFormatType.Expression, format.Type);
            Assert.Equal("=A1>B1", format.Formula1);
        }

        [Fact]
        public void ConditionalFormatStyle_CompleteStyle()
        {
            var style = new ConditionalFormatStyle
            {
                FontColor = ChartColor.Red,
                Bold = true,
                Italic = true,
                FillColor = ChartColor.Yellow,
                NumberFormat = "0.00%",
                Border = new ConditionalFormatBorder
                {
                    TopColor = ChartColor.Black,
                    BottomColor = ChartColor.Black,
                    LeftColor = ChartColor.Black,
                    RightColor = ChartColor.Black
                }
            };

            Assert.Equal(ChartColor.Red, style.FontColor);
            Assert.True(style.Bold);
            Assert.True(style.Italic);
            Assert.Equal(ChartColor.Yellow, style.FillColor);
            Assert.Equal("0.00%", style.NumberFormat);
            Assert.NotNull(style.Border);
            Assert.Equal(ChartColor.Black, style.Border.TopColor);
        }

        [Fact]
        public void ColorScale_TwoColorScale()
        {
            var colorScale = new ColorScale
            {
                Minimum = new ColorScalePoint
                {
                    Type = ColorScaleValueType.MinValue,
                    Color = ChartColor.White
                },
                Maximum = new ColorScalePoint
                {
                    Type = ColorScaleValueType.MaxValue,
                    Color = ChartColor.Red
                }
            };

            Assert.Equal(ColorScaleValueType.MinValue, colorScale.Minimum.Type);
            Assert.Equal(ChartColor.White, colorScale.Minimum.Color);
            Assert.Equal(ColorScaleValueType.MaxValue, colorScale.Maximum.Type);
            Assert.Equal(ChartColor.Red, colorScale.Maximum.Color);
            Assert.Null(colorScale.Midpoint);
        }

        [Fact]
        public void ColorScale_ThreeColorScale()
        {
            var colorScale = new ColorScale
            {
                Minimum = new ColorScalePoint
                {
                    Type = ColorScaleValueType.MinValue,
                    Color = ChartColor.Green
                },
                Midpoint = new ColorScalePoint
                {
                    Type = ColorScaleValueType.Percentile,
                    Value = 50,
                    Color = ChartColor.Yellow
                },
                Maximum = new ColorScalePoint
                {
                    Type = ColorScaleValueType.MaxValue,
                    Color = ChartColor.Red
                }
            };

            Assert.NotNull(colorScale.Midpoint);
            Assert.Equal(ColorScaleValueType.Percentile, colorScale.Midpoint.Type);
            Assert.Equal(50, colorScale.Midpoint.Value);
            Assert.Equal(ChartColor.Yellow, colorScale.Midpoint.Color);
        }

        [Fact]
        public void ColorScale_NumericValues()
        {
            var colorScale = new ColorScale
            {
                Minimum = new ColorScalePoint
                {
                    Type = ColorScaleValueType.Num,
                    Value = 0,
                    Color = ChartColor.Blue
                },
                Maximum = new ColorScalePoint
                {
                    Type = ColorScaleValueType.Num,
                    Value = 100,
                    Color = ChartColor.Red
                }
            };

            Assert.Equal(ColorScaleValueType.Num, colorScale.Minimum.Type);
            Assert.Equal(0, colorScale.Minimum.Value);
            Assert.Equal(ColorScaleValueType.Num, colorScale.Maximum.Type);
            Assert.Equal(100, colorScale.Maximum.Value);
        }

        [Fact]
        public void DataBar_DefaultValues()
        {
            var dataBar = new DataBar();

            Assert.Equal(DataBarValueType.MinValue, dataBar.Minimum.Type);
            Assert.Equal(DataBarValueType.MaxValue, dataBar.Maximum.Type);
            Assert.Equal(ChartColor.Blue, dataBar.Color);
            Assert.True(dataBar.ShowValue);
            Assert.Equal(DataBarDirection.LeftToRight, dataBar.Direction);
            Assert.Equal(DataBarAxisPosition.Automatic, dataBar.AxisPosition);
        }

        [Fact]
        public void DataBar_CustomConfiguration()
        {
            var dataBar = new DataBar
            {
                Minimum = new DataBarPoint
                {
                    Type = DataBarValueType.Num,
                    Value = 0
                },
                Maximum = new DataBarPoint
                {
                    Type = DataBarValueType.Num,
                    Value = 100
                },
                Color = ChartColor.Green,
                ShowValue = false,
                BorderColor = ChartColor.DarkGreen,
                Direction = DataBarDirection.RightToLeft,
                AxisPosition = DataBarAxisPosition.Middle
            };

            Assert.Equal(DataBarValueType.Num, dataBar.Minimum.Type);
            Assert.Equal(0, dataBar.Minimum.Value);
            Assert.Equal(DataBarValueType.Num, dataBar.Maximum.Type);
            Assert.Equal(100, dataBar.Maximum.Value);
            Assert.Equal(ChartColor.Green, dataBar.Color);
            Assert.False(dataBar.ShowValue);
            Assert.Equal(ChartColor.DarkGreen, dataBar.BorderColor);
            Assert.Equal(DataBarDirection.RightToLeft, dataBar.Direction);
            Assert.Equal(DataBarAxisPosition.Middle, dataBar.AxisPosition);
        }

        [Fact]
        public void IconSet_DefaultValues()
        {
            var iconSet = new IconSet();

            Assert.Equal(IconSetType.ThreeTrafficLights, iconSet.Type);
            Assert.True(iconSet.ShowValue);
            Assert.False(iconSet.Reverse);
        }

        [Theory]
        [InlineData(IconSetType.ThreeArrows)]
        [InlineData(IconSetType.ThreeArrowsGray)]
        [InlineData(IconSetType.ThreeFlags)]
        [InlineData(IconSetType.ThreeTrafficLights)]
        [InlineData(IconSetType.ThreeSigns)]
        [InlineData(IconSetType.ThreeSymbols)]
        [InlineData(IconSetType.ThreeSymbols2)]
        [InlineData(IconSetType.FourArrows)]
        [InlineData(IconSetType.FourArrowsGray)]
        [InlineData(IconSetType.FourRedToBlack)]
        [InlineData(IconSetType.FourRatings)]
        [InlineData(IconSetType.FourTrafficLights)]
        [InlineData(IconSetType.FiveArrows)]
        [InlineData(IconSetType.FiveArrowsGray)]
        [InlineData(IconSetType.FiveRatings)]
        [InlineData(IconSetType.FiveQuarters)]
        [InlineData(IconSetType.ThreeStars)]
        [InlineData(IconSetType.ThreeTriangles)]
        [InlineData(IconSetType.FiveBoxes)]
        public void IconSet_AllTypes(IconSetType type)
        {
            var iconSet = new IconSet { Type = type };
            Assert.Equal(type, iconSet.Type);
        }

        [Fact]
        public void IconSet_WithThresholds()
        {
            var iconSet = new IconSet
            {
                Type = IconSetType.ThreeTrafficLights,
                Thresholds = new List<IconThreshold>
                {
                    new() { Type = IconValueType.Percent, Value = 0 },
                    new() { Type = IconValueType.Percent, Value = 33 },
                    new() { Type = IconValueType.Percent, Value = 67 }
                }
            };

            Assert.Equal(3, iconSet.Thresholds.Count);
            Assert.Equal(33, iconSet.Thresholds[1].Value);
        }

        [Fact]
        public void CellRange_FromCells()
        {
            var range = CellRange.FromCells(5, 3, 15, 8);

            Assert.Equal(5, range.FirstRow);
            Assert.Equal(3, range.FirstCol);
            Assert.Equal(15, range.LastRow);
            Assert.Equal(8, range.LastCol);
        }

        [Fact]
        public void ConditionalFormatData_WithColorScale()
        {
            var format = new ConditionalFormatData
            {
                Type = ConditionalFormatType.ColorScale,
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                },
                ColorScale = new ColorScale
                {
                    Minimum = new ColorScalePoint
                    {
                        Type = ColorScaleValueType.MinValue,
                        Color = ChartColor.Green
                    },
                    Midpoint = new ColorScalePoint
                    {
                        Type = ColorScaleValueType.Percentile,
                        Value = 50,
                        Color = ChartColor.Yellow
                    },
                    Maximum = new ColorScalePoint
                    {
                        Type = ColorScaleValueType.MaxValue,
                        Color = ChartColor.Red
                    }
                }
            };

            Assert.Equal(ConditionalFormatType.ColorScale, format.Type);
            Assert.NotNull(format.ColorScale);
            Assert.NotNull(format.ColorScale.Midpoint);
        }

        [Fact]
        public void ConditionalFormatData_WithDataBar()
        {
            var format = new ConditionalFormatData
            {
                Type = ConditionalFormatType.DataBar,
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                },
                DataBar = new DataBar
                {
                    Color = ChartColor.Blue,
                    ShowValue = true
                }
            };

            Assert.Equal(ConditionalFormatType.DataBar, format.Type);
            Assert.NotNull(format.DataBar);
            Assert.Equal(ChartColor.Blue, format.DataBar.Color);
        }

        [Fact]
        public void ConditionalFormatData_WithIconSet()
        {
            var format = new ConditionalFormatData
            {
                Type = ConditionalFormatType.IconSet,
                Ranges = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                },
                IconSet = new IconSet
                {
                    Type = IconSetType.ThreeArrows,
                    ShowValue = false,
                    Reverse = true
                }
            };

            Assert.Equal(ConditionalFormatType.IconSet, format.Type);
            Assert.NotNull(format.IconSet);
            Assert.Equal(IconSetType.ThreeArrows, format.IconSet.Type);
            Assert.False(format.IconSet.ShowValue);
            Assert.True(format.IconSet.Reverse);
        }

        [Fact]
        public void ConditionalFormatData_PriorityAndStopIfTrue()
        {
            var format = new ConditionalFormatData
            {
                Priority = 5,
                StopIfTrue = true
            };

            Assert.Equal(5, format.Priority);
            Assert.True(format.StopIfTrue);
        }

        // 条件格式写入器测试
        [Fact]
        public void ConditionalFormatWriter_CreatePooled()
        {
            var writer = ConditionalFormatWriter.CreatePooled(out var buffer, 8192);
            try
            {
                Assert.True(buffer.Length >= 8192);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void ConditionalFormatWriter_WritesCellIsRule()
        {
            var writer = ConditionalFormatWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var format = new ConditionalFormatData
                {
                    Type = ConditionalFormatType.CellIs,
                    Operator = ComparisonOperator.GreaterThan,
                    Formula1 = "100",
                    Ranges = new List<CellRange>
                    {
                        new() { FirstRow = 0, FirstCol = 0, LastRow = 10, LastCol = 5 }
                    },
                    Style = new ConditionalFormatStyle
                    {
                        FillColor = ChartColor.Red
                    }
                };

                var bytesWritten = writer.WriteConditionalFormat(format, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void ConditionalFormatWriter_WritesColorScale()
        {
            var writer = ConditionalFormatWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var format = new ConditionalFormatData
                {
                    Type = ConditionalFormatType.ColorScale,
                    Ranges = new List<CellRange>
                    {
                        new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                    },
                    ColorScale = new ColorScale
                    {
                        Minimum = new ColorScalePoint { Type = ColorScaleValueType.MinValue, Color = ChartColor.Green },
                        Maximum = new ColorScalePoint { Type = ColorScaleValueType.MaxValue, Color = ChartColor.Red }
                    }
                };

                var bytesWritten = writer.WriteConditionalFormat(format, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void ConditionalFormatWriter_WritesDataBar()
        {
            var writer = ConditionalFormatWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var format = new ConditionalFormatData
                {
                    Type = ConditionalFormatType.DataBar,
                    Ranges = new List<CellRange>
                    {
                        new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                    },
                    DataBar = new DataBar
                    {
                        Color = ChartColor.Blue,
                        ShowValue = true
                    }
                };

                var bytesWritten = writer.WriteConditionalFormat(format, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void ConditionalFormatWriter_WritesIconSet()
        {
            var writer = ConditionalFormatWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var format = new ConditionalFormatData
                {
                    Type = ConditionalFormatType.IconSet,
                    Ranges = new List<CellRange>
                    {
                        new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 0 }
                    },
                    IconSet = new IconSet
                    {
                        Type = IconSetType.ThreeTrafficLights,
                        ShowValue = true,
                        Thresholds = new List<IconThreshold>
                        {
                            new() { Type = IconValueType.Percent, Value = 0 },
                            new() { Type = IconValueType.Percent, Value = 33 },
                            new() { Type = IconValueType.Percent, Value = 67 }
                        }
                    }
                };

                var bytesWritten = writer.WriteConditionalFormat(format, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void ConditionalFormatWriter_WritesMultipleRanges()
        {
            var writer = ConditionalFormatWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var format = new ConditionalFormatData
                {
                    Type = ConditionalFormatType.CellIs,
                    Operator = ComparisonOperator.Between,
                    Formula1 = "0",
                    Formula2 = "100",
                    Ranges = new List<CellRange>
                    {
                        new() { FirstRow = 0, FirstCol = 0, LastRow = 10, LastCol = 5 },
                        new() { FirstRow = 20, FirstCol = 0, LastRow = 30, LastCol = 5 },
                        new() { FirstRow = 40, FirstCol = 0, LastRow = 50, LastCol = 5 }
                    },
                    Style = new ConditionalFormatStyle
                    {
                        FillColor = ChartColor.Yellow
                    }
                };

                var bytesWritten = writer.WriteConditionalFormat(format, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void ConditionalFormatWriter_WritesCompleteStyle()
        {
            var writer = ConditionalFormatWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var format = new ConditionalFormatData
                {
                    Type = ConditionalFormatType.Expression,
                    Formula1 = "=A1>B1",
                    Ranges = new List<CellRange>
                    {
                        new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 10 }
                    },
                    Style = new ConditionalFormatStyle
                    {
                        FontColor = ChartColor.Red,
                        Bold = true,
                        Italic = true,
                        FillColor = ChartColor.Yellow,
                        NumberFormat = "0.00%",
                        Border = new ConditionalFormatBorder
                        {
                            TopColor = ChartColor.Black,
                            BottomColor = ChartColor.Black
                        }
                    }
                };

                var bytesWritten = writer.WriteConditionalFormat(format, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        // 枚举值测试
        [Theory]
        [InlineData(ConditionalFormatType.CellIs, 0)]
        [InlineData(ConditionalFormatType.ContainsText, 1)]
        [InlineData(ConditionalFormatType.NotContainsText, 2)]
        [InlineData(ConditionalFormatType.BeginsWith, 3)]
        [InlineData(ConditionalFormatType.EndsWith, 4)]
        [InlineData(ConditionalFormatType.ContainsDate, 5)]
        [InlineData(ConditionalFormatType.Top10, 6)]
        [InlineData(ConditionalFormatType.UniqueValues, 7)]
        [InlineData(ConditionalFormatType.Expression, 8)]
        [InlineData(ConditionalFormatType.ColorScale, 9)]
        [InlineData(ConditionalFormatType.DataBar, 10)]
        [InlineData(ConditionalFormatType.IconSet, 11)]
        [InlineData(ConditionalFormatType.AboveAverage, 12)]
        [InlineData(ConditionalFormatType.BelowAverage, 13)]
        public void ConditionalFormatType_HasCorrectValues(ConditionalFormatType type, byte expected)
        {
            Assert.Equal(expected, (byte)type);
        }

        [Theory]
        [InlineData(ComparisonOperator.None, 0)]
        [InlineData(ComparisonOperator.Between, 1)]
        [InlineData(ComparisonOperator.NotBetween, 2)]
        [InlineData(ComparisonOperator.Equal, 3)]
        [InlineData(ComparisonOperator.NotEqual, 4)]
        [InlineData(ComparisonOperator.GreaterThan, 5)]
        [InlineData(ComparisonOperator.LessThan, 6)]
        [InlineData(ComparisonOperator.GreaterThanOrEqual, 7)]
        [InlineData(ComparisonOperator.LessThanOrEqual, 8)]
        public void ComparisonOperator_HasCorrectValues(ComparisonOperator op, byte expected)
        {
            Assert.Equal(expected, (byte)op);
        }

        [Theory]
        [InlineData(ColorScaleValueType.Num, 0)]
        [InlineData(ColorScaleValueType.MinValue, 1)]
        [InlineData(ColorScaleValueType.MaxValue, 2)]
        [InlineData(ColorScaleValueType.Percent, 3)]
        [InlineData(ColorScaleValueType.Percentile, 4)]
        [InlineData(ColorScaleValueType.Formula, 5)]
        public void ColorScaleValueType_HasCorrectValues(ColorScaleValueType type, byte expected)
        {
            Assert.Equal(expected, (byte)type);
        }

        [Theory]
        [InlineData(DataBarValueType.Num, 0)]
        [InlineData(DataBarValueType.MinValue, 1)]
        [InlineData(DataBarValueType.MaxValue, 2)]
        [InlineData(DataBarValueType.Percentile, 3)]
        [InlineData(DataBarValueType.Formula, 4)]
        [InlineData(DataBarValueType.Auto, 5)]
        public void DataBarValueType_HasCorrectValues(DataBarValueType type, byte expected)
        {
            Assert.Equal(expected, (byte)type);
        }

        [Theory]
        [InlineData(DataBarDirection.LeftToRight, 0)]
        [InlineData(DataBarDirection.RightToLeft, 1)]
        public void DataBarDirection_HasCorrectValues(DataBarDirection dir, byte expected)
        {
            Assert.Equal(expected, (byte)dir);
        }

        [Theory]
        [InlineData(DataBarAxisPosition.Automatic, 0)]
        [InlineData(DataBarAxisPosition.Middle, 1)]
        [InlineData(DataBarAxisPosition.None, 2)]
        public void DataBarAxisPosition_HasCorrectValues(DataBarAxisPosition pos, byte expected)
        {
            Assert.Equal(expected, (byte)pos);
        }

        [Theory]
        [InlineData(IconValueType.Num, 0)]
        [InlineData(IconValueType.Percent, 1)]
        [InlineData(IconValueType.Formula, 2)]
        [InlineData(IconValueType.Percentile, 3)]
        public void IconValueType_HasCorrectValues(IconValueType type, byte expected)
        {
            Assert.Equal(expected, (byte)type);
        }

        [Theory]
        [InlineData(IconOperator.GreaterThan, 0)]
        [InlineData(IconOperator.GreaterThanOrEqual, 1)]
        public void IconOperator_HasCorrectValues(IconOperator op, byte expected)
        {
            Assert.Equal(expected, (byte)op);
        }
    }
}
