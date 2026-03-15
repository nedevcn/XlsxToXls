using System.Collections.Generic;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    public class PageSetupTests
    {
        [Fact]
        public void PageSetupData_DefaultValues()
        {
            var setup = new PageSetupData();

            Assert.Equal(PaperSize.Letter, setup.PaperSize);
            Assert.Equal(PageOrientation.Portrait, setup.Orientation);
            Assert.Null(setup.Scale);
            Assert.Null(setup.FitToWidth);
            Assert.Null(setup.FitToHeight);
            Assert.NotNull(setup.Margins);
            Assert.NotNull(setup.PrintArea);
            Assert.Empty(setup.PrintArea);
            Assert.Null(setup.PrintTitleRows);
            Assert.Null(setup.PrintTitleColumns);
            Assert.Null(setup.Header);
            Assert.Null(setup.Footer);
            Assert.Equal(0.5, setup.HeaderMargin);
            Assert.Equal(0.5, setup.FooterMargin);
            Assert.Equal(1, setup.FirstPageNumber);
            Assert.Equal(600, setup.PrintQuality);
            Assert.Null(setup.StartPageNumber);
            Assert.False(setup.CenterHorizontally);
            Assert.False(setup.CenterVertically);
            Assert.Equal(1, setup.Copies);
            Assert.True(setup.PrintGridlines);
            Assert.False(setup.PrintHeadings);
            Assert.False(setup.BlackAndWhite);
            Assert.Equal(PrintComments.None, setup.PrintComments);
            Assert.Equal(PageOrder.DownThenOver, setup.PageOrder);
            Assert.Equal(CellErrorPrint.Displayed, setup.CellErrors);
            Assert.False(setup.DraftQuality);
        }

        [Fact]
        public void PageSetupData_LandscapeOrientation()
        {
            var setup = new PageSetupData
            {
                Orientation = PageOrientation.Landscape,
                PaperSize = PaperSize.A4
            };

            Assert.Equal(PageOrientation.Landscape, setup.Orientation);
            Assert.Equal(PaperSize.A4, setup.PaperSize);
        }

        [Fact]
        public void PageSetupData_ScaleFactor()
        {
            var setup = new PageSetupData
            {
                Scale = 150
            };

            Assert.Equal(150, setup.Scale);
        }

        [Fact]
        public void PageSetupData_FitToPages()
        {
            var setup = new PageSetupData
            {
                FitToWidth = 2,
                FitToHeight = 3
            };

            Assert.Equal(2, setup.FitToWidth);
            Assert.Equal(3, setup.FitToHeight);
            Assert.Null(setup.Scale);
        }

        [Fact]
        public void PageMargins_DefaultValues()
        {
            var margins = new PageMargins();

            Assert.Equal(0.75, margins.Left);
            Assert.Equal(0.75, margins.Right);
            Assert.Equal(1.0, margins.Top);
            Assert.Equal(1.0, margins.Bottom);
            Assert.Equal(0.5, margins.Header);
            Assert.Equal(0.5, margins.Footer);
        }

        [Fact]
        public void PageMargins_CustomValues()
        {
            var margins = new PageMargins
            {
                Left = 1.0,
                Right = 1.0,
                Top = 1.5,
                Bottom = 1.5,
                Header = 0.75,
                Footer = 0.75
            };

            Assert.Equal(1.0, margins.Left);
            Assert.Equal(1.0, margins.Right);
            Assert.Equal(1.5, margins.Top);
            Assert.Equal(1.5, margins.Bottom);
            Assert.Equal(0.75, margins.Header);
            Assert.Equal(0.75, margins.Footer);
        }

        [Fact]
        public void PageSetupData_PrintArea()
        {
            var setup = new PageSetupData
            {
                PrintArea = new List<CellRange>
                {
                    new() { FirstRow = 0, FirstCol = 0, LastRow = 10, LastCol = 5 },
                    new() { FirstRow = 20, FirstCol = 0, LastRow = 30, LastCol = 5 }
                }
            };

            Assert.Equal(2, setup.PrintArea.Count);
            Assert.Equal(0, setup.PrintArea[0].FirstRow);
            Assert.Equal(20, setup.PrintArea[1].FirstRow);
        }

        [Fact]
        public void PageSetupData_PrintTitleRows()
        {
            var setup = new PageSetupData
            {
                PrintTitleRows = new CellRange
                {
                    FirstRow = 0,
                    LastRow = 1,
                    FirstCol = 0,
                    LastCol = 255
                }
            };

            Assert.NotNull(setup.PrintTitleRows);
            Assert.Equal(0, setup.PrintTitleRows.FirstRow);
            Assert.Equal(1, setup.PrintTitleRows.LastRow);
        }

        [Fact]
        public void PageSetupData_PrintTitleColumns()
        {
            var setup = new PageSetupData
            {
                PrintTitleColumns = new CellRange
                {
                    FirstRow = 0,
                    LastRow = 65535,
                    FirstCol = 0,
                    LastCol = 2
                }
            };

            Assert.NotNull(setup.PrintTitleColumns);
            Assert.Equal(0, setup.PrintTitleColumns.FirstCol);
            Assert.Equal(2, setup.PrintTitleColumns.LastCol);
        }

        [Fact]
        public void PageSetupData_HeaderFooter()
        {
            var setup = new PageSetupData
            {
                Header = "&CPage &P of &N",
                Footer = "&LCompany Name&CConfidential&R&D",
                HeaderMargin = 0.75,
                FooterMargin = 0.75
            };

            Assert.Equal("&CPage &P of &N", setup.Header);
            Assert.Equal("&LCompany Name&CConfidential&R&D", setup.Footer);
            Assert.Equal(0.75, setup.HeaderMargin);
            Assert.Equal(0.75, setup.FooterMargin);
        }

        [Fact]
        public void PageSetupData_PrintOptions()
        {
            var setup = new PageSetupData
            {
                PrintGridlines = false,
                PrintHeadings = true,
                BlackAndWhite = true,
                DraftQuality = true,
                PrintComments = PrintComments.AtEnd,
                PageOrder = PageOrder.OverThenDown,
                CellErrors = CellErrorPrint.Blank,
                CenterHorizontally = true,
                CenterVertically = true
            };

            Assert.False(setup.PrintGridlines);
            Assert.True(setup.PrintHeadings);
            Assert.True(setup.BlackAndWhite);
            Assert.True(setup.DraftQuality);
            Assert.Equal(PrintComments.AtEnd, setup.PrintComments);
            Assert.Equal(PageOrder.OverThenDown, setup.PageOrder);
            Assert.Equal(CellErrorPrint.Blank, setup.CellErrors);
            Assert.True(setup.CenterHorizontally);
            Assert.True(setup.CenterVertically);
        }

        [Fact]
        public void PageSetupData_Copies()
        {
            var setup = new PageSetupData
            {
                Copies = 3
            };

            Assert.Equal(3, setup.Copies);
        }

        [Fact]
        public void PageSetupData_FirstPageNumber()
        {
            var setup = new PageSetupData
            {
                FirstPageNumber = 5
            };

            Assert.Equal(5, setup.FirstPageNumber);
        }

        [Fact]
        public void PageSetupData_PrintQuality()
        {
            var setup = new PageSetupData
            {
                PrintQuality = 300
            };

            Assert.Equal(300, setup.PrintQuality);
        }

        // Paper size tests
        [Theory]
        [InlineData(PaperSize.Letter, 1)]
        [InlineData(PaperSize.LetterSmall, 2)]
        [InlineData(PaperSize.Tabloid, 3)]
        [InlineData(PaperSize.Ledger, 4)]
        [InlineData(PaperSize.Legal, 5)]
        [InlineData(PaperSize.Statement, 6)]
        [InlineData(PaperSize.Executive, 7)]
        [InlineData(PaperSize.A3, 8)]
        [InlineData(PaperSize.A4, 9)]
        [InlineData(PaperSize.A4Small, 10)]
        [InlineData(PaperSize.A5, 11)]
        [InlineData(PaperSize.B4, 12)]
        [InlineData(PaperSize.B5, 13)]
        [InlineData(PaperSize.Folio, 14)]
        [InlineData(PaperSize.Quarto, 15)]
        public void PaperSize_HasCorrectValues(PaperSize size, ushort expected)
        {
            Assert.Equal(expected, (ushort)size);
        }

        [Theory]
        [InlineData(PageOrientation.Portrait, 0)]
        [InlineData(PageOrientation.Landscape, 1)]
        public void PageOrientation_HasCorrectValues(PageOrientation orientation, byte expected)
        {
            Assert.Equal(expected, (byte)orientation);
        }

        [Theory]
        [InlineData(PageOrder.DownThenOver, 0)]
        [InlineData(PageOrder.OverThenDown, 1)]
        public void PageOrder_HasCorrectValues(PageOrder order, byte expected)
        {
            Assert.Equal(expected, (byte)order);
        }

        [Theory]
        [InlineData(PrintComments.None, 0)]
        [InlineData(PrintComments.AtEnd, 1)]
        [InlineData(PrintComments.AsDisplayed, 2)]
        public void PrintComments_HasCorrectValues(PrintComments comments, byte expected)
        {
            Assert.Equal(expected, (byte)comments);
        }

        [Theory]
        [InlineData(CellErrorPrint.Displayed, 0)]
        [InlineData(CellErrorPrint.Blank, 1)]
        [InlineData(CellErrorPrint.DashDash, 2)]
        [InlineData(CellErrorPrint.NA, 3)]
        public void CellErrorPrint_HasCorrectValues(CellErrorPrint error, byte expected)
        {
            Assert.Equal(expected, (byte)error);
        }

        // Header/Footer codes tests
        [Fact]
        public void HeaderFooterCodes_Left()
        {
            Assert.Equal("&L", HeaderFooterCodes.Left);
        }

        [Fact]
        public void HeaderFooterCodes_Center()
        {
            Assert.Equal("&C", HeaderFooterCodes.Center);
        }

        [Fact]
        public void HeaderFooterCodes_Right()
        {
            Assert.Equal("&R", HeaderFooterCodes.Right);
        }

        [Fact]
        public void HeaderFooterCodes_PageNumber()
        {
            Assert.Equal("&P", HeaderFooterCodes.PageNumber);
        }

        [Fact]
        public void HeaderFooterCodes_TotalPages()
        {
            Assert.Equal("&N", HeaderFooterCodes.TotalPages);
        }

        [Fact]
        public void HeaderFooterCodes_Date()
        {
            Assert.Equal("&D", HeaderFooterCodes.Date);
        }

        [Fact]
        public void HeaderFooterCodes_Time()
        {
            Assert.Equal("&T", HeaderFooterCodes.Time);
        }

        [Fact]
        public void HeaderFooterCodes_FileName()
        {
            Assert.Equal("&F", HeaderFooterCodes.FileName);
        }

        [Fact]
        public void HeaderFooterCodes_SheetName()
        {
            Assert.Equal("&A", HeaderFooterCodes.SheetName);
        }

        [Fact]
        public void HeaderFooterCodes_Bold()
        {
            Assert.Equal("&B", HeaderFooterCodes.Bold);
        }

        [Fact]
        public void HeaderFooterCodes_Italic()
        {
            Assert.Equal("&I", HeaderFooterCodes.Italic);
        }

        [Fact]
        public void HeaderFooterCodes_Underline()
        {
            Assert.Equal("&U", HeaderFooterCodes.Underline);
        }

        [Fact]
        public void HeaderFooterCodes_Strikethrough()
        {
            Assert.Equal("&S", HeaderFooterCodes.Strikethrough);
        }

        [Fact]
        public void HeaderFooterCodes_FontSize()
        {
            Assert.Equal("&\"12\"", HeaderFooterCodes.FontSize(12));
        }

        [Fact]
        public void HeaderFooterCodes_FontName()
        {
            Assert.Equal("&\"Arial\"", HeaderFooterCodes.FontName("Arial"));
        }

        // PageSetupWriter tests
        [Fact]
        public void PageSetupWriter_CreatePooled()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 4096);
            try
            {
                Assert.True(buffer.Length >= 4096);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesDefaultSetup()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData();
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesLandscapeOrientation()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData
                {
                    Orientation = PageOrientation.Landscape,
                    PaperSize = PaperSize.A4
                };
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesCustomMargins()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData
                {
                    Margins = new PageMargins
                    {
                        Left = 1.0,
                        Right = 1.0,
                        Top = 1.5,
                        Bottom = 1.5
                    }
                };
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesHeaderFooter()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData
                {
                    Header = "&CPage &P of &N",
                    Footer = "&LCompany Name&CConfidential"
                };
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesPrintArea()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData
                {
                    PrintArea = new List<CellRange>
                    {
                        new() { FirstRow = 0, FirstCol = 0, LastRow = 10, LastCol = 5 }
                    }
                };
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesMultiplePrintAreas()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData
                {
                    PrintArea = new List<CellRange>
                    {
                        new() { FirstRow = 0, FirstCol = 0, LastRow = 10, LastCol = 5 },
                        new() { FirstRow = 20, FirstCol = 0, LastRow = 30, LastCol = 5 }
                    }
                };
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesPrintTitleRows()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData
                {
                    PrintTitleRows = new CellRange
                    {
                        FirstRow = 0,
                        LastRow = 1,
                        FirstCol = 0,
                        LastCol = 255
                    }
                };
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesPrintTitleColumns()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData
                {
                    PrintTitleColumns = new CellRange
                    {
                        FirstRow = 0,
                        LastRow = 65535,
                        FirstCol = 0,
                        LastCol = 2
                    }
                };
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesBothPrintTitles()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData
                {
                    PrintTitleRows = new CellRange
                    {
                        FirstRow = 0,
                        LastRow = 1,
                        FirstCol = 0,
                        LastCol = 255
                    },
                    PrintTitleColumns = new CellRange
                    {
                        FirstRow = 0,
                        LastRow = 65535,
                        FirstCol = 0,
                        LastCol = 2
                    }
                };
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesScaleFactor()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData
                {
                    Scale = 150
                };
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesFitToPages()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData
                {
                    FitToWidth = 2,
                    FitToHeight = 3
                };
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesPrintOptions()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData
                {
                    PrintGridlines = false,
                    PrintHeadings = true,
                    CenterHorizontally = true,
                    CenterVertically = true,
                    BlackAndWhite = true,
                    DraftQuality = true
                };
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesCopies()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 8192);
            try
            {
                var setup = new PageSetupData
                {
                    Copies = 5
                };
                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }

        [Fact]
        public void PageSetupWriter_WritesCompleteSetup()
        {
            var writer = PageSetupWriter.CreatePooled(out var buffer, 16384);
            try
            {
                var setup = new PageSetupData
                {
                    PaperSize = PaperSize.A4,
                    Orientation = PageOrientation.Landscape,
                    Scale = 100,
                    Margins = new PageMargins
                    {
                        Left = 0.5,
                        Right = 0.5,
                        Top = 0.75,
                        Bottom = 0.75
                    },
                    PrintArea = new List<CellRange>
                    {
                        new() { FirstRow = 0, FirstCol = 0, LastRow = 100, LastCol = 20 }
                    },
                    PrintTitleRows = new CellRange
                    {
                        FirstRow = 0,
                        LastRow = 2,
                        FirstCol = 0,
                        LastCol = 255
                    },
                    Header = "&CPage &P of &N",
                    Footer = "&LCompany Name&CConfidential&R&D",
                    PrintGridlines = true,
                    PrintHeadings = false,
                    CenterHorizontally = true,
                    CenterVertically = false,
                    Copies = 1
                };

                var bytesWritten = writer.WritePageSetup(setup, 0);
                Assert.True(bytesWritten > 0);
            }
            finally
            {
                writer.Dispose();
            }
        }
    }
}
