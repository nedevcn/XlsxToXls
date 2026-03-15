using System;
using System.Buffers;
using System.Collections.Generic;
using System.IO;
using Nedev.FileConverters.XlsxToXls.Internal;
using Xunit;

namespace Nedev.FileConverters.XlsxToXls.Tests
{
    public class ShapeTests
    {
        [Fact]
        public void ShapeData_DefaultValues()
        {
            var shape = new ShapeData();

            Assert.Equal(0, shape.ShapeId);
            // Default ShapeType is Picture (0) as it's the first enum value
            Assert.Equal(ShapeType.Picture, shape.Type);
            Assert.Empty(shape.Name);
            Assert.NotNull(shape.Position);
            Assert.True(shape.Visible);
            Assert.Equal(0, shape.Rotation);
            Assert.Equal(0, shape.ZOrder);
        }

        [Fact]
        public void ShapeData_WithValues()
        {
            var shape = new ShapeData
            {
                ShapeId = 1,
                Type = ShapeType.Picture,
                Name = "TestPicture",
                Description = "A test picture",
                Visible = false,
                Rotation = 45,
                ZOrder = 2
            };

            Assert.Equal(1, shape.ShapeId);
            Assert.Equal(ShapeType.Picture, shape.Type);
            Assert.Equal("TestPicture", shape.Name);
            Assert.Equal("A test picture", shape.Description);
            Assert.False(shape.Visible);
            Assert.Equal(45, shape.Rotation);
            Assert.Equal(2, shape.ZOrder);
        }

        [Fact]
        public void ShapePosition_DefaultValues()
        {
            var position = new ShapePosition();

            Assert.Equal(AnchorType.OneCell, position.AnchorType);
            Assert.Equal(0, position.FromCol);
            Assert.Equal(0, position.FromRow);
            Assert.Equal(0, position.FromColOffset);
            Assert.Equal(0, position.FromRowOffset);
            Assert.Equal(0, position.WidthEmu);
            Assert.Equal(0, position.HeightEmu);
        }

        [Fact]
        public void ShapePosition_WithValues()
        {
            var position = new ShapePosition
            {
                AnchorType = AnchorType.TwoCell,
                FromCol = 1,
                FromRow = 2,
                FromColOffset = 100000,
                FromRowOffset = 200000,
                ToCol = 5,
                ToRow = 10,
                ToColOffset = 50000,
                ToRowOffset = 100000,
                WidthEmu = 1000000,
                HeightEmu = 500000
            };

            Assert.Equal(AnchorType.TwoCell, position.AnchorType);
            Assert.Equal(1, position.FromCol);
            Assert.Equal(2, position.FromRow);
            Assert.Equal(100000, position.FromColOffset);
            Assert.Equal(200000, position.FromRowOffset);
            Assert.Equal(5, position.ToCol);
            Assert.Equal(10, position.ToRow);
            Assert.Equal(50000, position.ToColOffset);
            Assert.Equal(100000, position.ToRowOffset);
            Assert.Equal(1000000, position.WidthEmu);
            Assert.Equal(500000, position.HeightEmu);
        }

        [Fact]
        public void ImageData_DefaultValues()
        {
            var image = new ImageData();

            Assert.Equal(ImageFormat.Unknown, image.Format);
            Assert.Empty(image.Data);
            Assert.Null(image.RelationshipId);
            Assert.Null(image.ContentType);
        }

        [Fact]
        public void ImageData_WithValues()
        {
            var imageData = new byte[] { 0x89, 0x50, 0x4E, 0x47 }; // PNG signature
            var image = new ImageData
            {
                Format = ImageFormat.Png,
                Data = imageData,
                RelationshipId = "rId1",
                ContentType = "image/png"
            };

            Assert.Equal(ImageFormat.Png, image.Format);
            Assert.Equal(imageData, image.Data);
            Assert.Equal("rId1", image.RelationshipId);
            Assert.Equal("image/png", image.ContentType);
        }

        [Theory]
        [InlineData(ShapeType.Picture)]
        [InlineData(ShapeType.Rectangle)]
        [InlineData(ShapeType.Oval)]
        [InlineData(ShapeType.Line)]
        [InlineData(ShapeType.TextBox)]
        public void ShapeType_AllTypes(ShapeType type)
        {
            var shape = new ShapeData { Type = type };
            Assert.Equal(type, shape.Type);
        }

        [Theory]
        [InlineData(ImageFormat.Png)]
        [InlineData(ImageFormat.Jpeg)]
        [InlineData(ImageFormat.Gif)]
        [InlineData(ImageFormat.Bmp)]
        [InlineData(ImageFormat.Tiff)]
        [InlineData(ImageFormat.Wmf)]
        [InlineData(ImageFormat.Emf)]
        public void ImageFormat_AllFormats(ImageFormat format)
        {
            var image = new ImageData { Format = format };
            Assert.Equal(format, image.Format);
        }

        [Fact]
        public void FillProperties_DefaultValues()
        {
            var fill = new FillProperties();

            Assert.Equal(FillType.Solid, fill.Type);
            Assert.Equal(0xFFFFFFFFu, fill.Color);
            Assert.Equal(0, fill.Transparency);
        }

        [Fact]
        public void LineProperties_DefaultValues()
        {
            var line = new LineProperties();

            Assert.Equal(ShapeLineStyle.Solid, line.Style);
            Assert.Equal(9525, line.Width); // 0.75pt default
            Assert.Equal(0xFF000000u, line.Color);
        }

        [Fact]
        public void TextProperties_DefaultValues()
        {
            var text = new TextProperties();

            Assert.Equal("Arial", text.FontName);
            Assert.Equal(11, text.FontSize);
            Assert.False(text.Bold);
            Assert.False(text.Italic);
            Assert.Equal(0xFF000000u, text.Color);
            Assert.Equal(TextAlignment.Left, text.HorizontalAlignment);
            Assert.Equal(TextVerticalAlignment.Top, text.VerticalAlignment);
            Assert.True(text.WrapText);
        }

        [Fact]
        public void ShapeInfo_RecordCreation()
        {
            var position = new ShapePositionInfo(
                AnchorType.OneCell, 0, 0, 0, 0, 0, 0, 0, 0, 1000000, 500000);

            var info = new ShapeInfo(
                1,
                ShapeType.Rectangle,
                "TestShape",
                position,
                null,
                true,
                0,
                1);

            Assert.Equal(1, info.ShapeId);
            Assert.Equal(ShapeType.Rectangle, info.Type);
            Assert.Equal("TestShape", info.Name);
            Assert.True(info.Visible);
            Assert.Equal(1, info.ZOrder);
        }

        [Fact]
        public void ShapeInfo_WithImage()
        {
            var imageData = new byte[] { 0xFF, 0xD8, 0xFF }; // JPEG signature
            var image = new ImageInfo(ImageFormat.Jpeg, imageData, "rId2", "image/jpeg");
            var position = new ShapePositionInfo(
                AnchorType.TwoCell, 1, 1, 0, 0, 5, 5, 0, 0, 2000000, 1500000);

            var info = new ShapeInfo(
                2,
                ShapeType.Picture,
                "TestPicture",
                position,
                image,
                true,
                0,
                2);

            Assert.Equal(ShapeType.Picture, info.Type);
            Assert.NotNull(info.Image);
            Assert.Equal(ImageFormat.Jpeg, info.Image.Value.Format);
            Assert.Equal(imageData, info.Image.Value.Data);
        }

        [Fact]
        public void AnchorType_AllTypes()
        {
            Assert.Equal(0, (int)AnchorType.Absolute);
            Assert.Equal(1, (int)AnchorType.OneCell);
            Assert.Equal(2, (int)AnchorType.TwoCell);
        }

        [Fact]
        public void TextAlignment_AllValues()
        {
            Assert.Equal(0, (int)TextAlignment.Left);
            Assert.Equal(1, (int)TextAlignment.Center);
            Assert.Equal(2, (int)TextAlignment.Right);
            Assert.Equal(3, (int)TextAlignment.Justify);
        }

        [Fact]
        public void TextVerticalAlignment_AllValues()
        {
            Assert.Equal(0, (int)TextVerticalAlignment.Top);
            Assert.Equal(1, (int)TextVerticalAlignment.Middle);
            Assert.Equal(2, (int)TextVerticalAlignment.Bottom);
        }

        [Fact]
        public void FillType_AllValues()
        {
            Assert.Equal(0, (int)FillType.None);
            Assert.Equal(1, (int)FillType.Solid);
            Assert.Equal(2, (int)FillType.Gradient);
            Assert.Equal(3, (int)FillType.Pattern);
            Assert.Equal(4, (int)FillType.Picture);
        }

        [Fact]
        public void ShapeLineStyle_AllValues()
        {
            Assert.Equal(0, (int)ShapeLineStyle.None);
            Assert.Equal(1, (int)ShapeLineStyle.Solid);
            Assert.Equal(2, (int)ShapeLineStyle.Dash);
            Assert.Equal(3, (int)ShapeLineStyle.Dot);
            Assert.Equal(4, (int)ShapeLineStyle.DashDot);
            Assert.Equal(5, (int)ShapeLineStyle.DashDotDot);
        }

        [Fact]
        public void ShapeWriter_CreatePooled()
        {
            var writer = ShapeWriter.CreatePooled(out var buffer, 1024);

            Assert.NotNull(buffer);
            Assert.True(buffer.Length >= 1024);

            writer.Dispose();
            ArrayPool<byte>.Shared.Return(buffer);
        }

        [Fact]
        public void ShapeWriter_WriteEmptyShapes()
        {
            var writer = ShapeWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var shapes = new List<ShapeInfo>();
                var position = writer.WriteAllShapes(shapes, 1);

                Assert.Equal(0, position);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        [Fact]
        public void ShapeWriter_WriteSingleShape()
        {
            var writer = ShapeWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var position = new ShapePositionInfo(
                    AnchorType.OneCell, 0, 0, 0, 0, 0, 0, 0, 0, 1000000, 500000);
                var shape = new ShapeInfo(
                    1, ShapeType.Rectangle, "Test", position, null, true, 0, 1);
                var shapes = new List<ShapeInfo> { shape };

                var written = writer.WriteAllShapes(shapes, 1);

                Assert.True(written > 0);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        [Fact]
        public void ShapeWriter_WriteMultipleShapes()
        {
            var writer = ShapeWriter.CreatePooled(out var buffer, 65536);
            try
            {
                var shapes = new List<ShapeInfo>();
                for (int i = 1; i <= 3; i++)
                {
                    var position = new ShapePositionInfo(
                        AnchorType.OneCell, i, i, 0, 0, 0, 0, 0, 0, 500000, 500000);
                    shapes.Add(new ShapeInfo(
                        i, ShapeType.Rectangle, $"Shape{i}", position, null, true, 0, i));
                }

                var written = writer.WriteAllShapes(shapes, 1);

                Assert.True(written > 0);
            }
            finally
            {
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }

        [Fact]
        public void ShapePositionInfo_AdditionalConstructor()
        {
            // Test the 13-parameter constructor
            var position = new ShapePositionInfo(
                AnchorType.TwoCell,
                1, 2, 100, 200,
                5, 10, 50, 100,
                1000000, 500000,
                0, 0); // Additional unused parameters

            Assert.Equal(AnchorType.TwoCell, position.AnchorType);
            Assert.Equal(1, position.FromCol);
            Assert.Equal(2, position.FromRow);
            Assert.Equal(100, position.FromColOffset);
            Assert.Equal(200, position.FromRowOffset);
            Assert.Equal(5, position.ToCol);
            Assert.Equal(10, position.ToRow);
            Assert.Equal(50, position.ToColOffset);
            Assert.Equal(100, position.ToRowOffset);
            Assert.Equal(1000000, position.WidthEmu);
            Assert.Equal(500000, position.HeightEmu);
        }
    }
}
