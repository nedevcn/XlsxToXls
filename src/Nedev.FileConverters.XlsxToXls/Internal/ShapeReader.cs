using System.IO.Compression;
using System.Xml;

namespace Nedev.FileConverters.XlsxToXls.Internal;

/// <summary>
/// Reads shapes and images from XLSX files.
/// </summary>
internal static class ShapeReader
{
    /// <summary>
    /// Reads all shapes from the worksheet drawing.
    /// </summary>
    public static List<ShapeInfo> ReadShapes(ZipArchive archive, string worksheetPath, Action<string>? log)
    {
        var shapes = new List<ShapeInfo>();
        var drawingPath = GetDrawingPath(archive, worksheetPath);

        if (string.IsNullOrEmpty(drawingPath))
        {
            log?.Invoke($"[ShapeReader] No drawing found for worksheet: {worksheetPath}");
            return shapes;
        }

        try
        {
            var drawingEntry = archive.GetEntry(drawingPath);
            if (drawingEntry == null)
            {
                log?.Invoke($"[ShapeReader] Drawing entry not found: {drawingPath}");
                return shapes;
            }

            using var stream = drawingEntry.Open();
            using var reader = XmlReader.Create(stream);
            var shapeId = 1;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "pic": // Picture
                            var picture = ReadPicture(reader, archive, drawingPath, shapeId, log);
                            if (picture.HasValue)
                            {
                                shapes.Add(picture.Value);
                                shapeId++;
                            }
                            break;

                        case "sp": // Shape
                            var shape = ReadShape(reader, shapeId, log);
                            if (shape.HasValue)
                            {
                                shapes.Add(shape.Value);
                                shapeId++;
                            }
                            break;

                        case "cxnSp": // Connector shape (line/arrow)
                            var connector = ReadConnector(reader, shapeId, log);
                            if (connector.HasValue)
                            {
                                shapes.Add(connector.Value);
                                shapeId++;
                            }
                            break;

                        case "grpSp": // Group shape
                            // Group shapes are complex - for now, skip or process children
                            log?.Invoke($"[ShapeReader] Group shape found but not fully supported");
                            break;
                    }
                }
            }

            log?.Invoke($"[ShapeReader] Read {shapes.Count} shapes from drawing");
        }
        catch (Exception ex)
        {
            log?.Invoke($"[ShapeReader] Error reading shapes: {ex.Message}");
        }

        return shapes;
    }

    /// <summary>
    /// Gets the drawing path from the worksheet relationships.
    /// </summary>
    private static string? GetDrawingPath(ZipArchive archive, string worksheetPath)
    {
        var relsPath = $"xl/worksheets/_rels/{Path.GetFileName(worksheetPath)}.rels";
        var relsEntry = archive.GetEntry(relsPath);

        if (relsEntry == null)
            return null;

        using var stream = relsEntry.Open();
        using var reader = XmlReader.Create(stream);

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship")
            {
                var type = reader.GetAttribute("Type");
                if (type?.Contains("drawing") == true)
                {
                    var target = reader.GetAttribute("Target");
                    if (!string.IsNullOrEmpty(target))
                    {
                        // Convert relative path to absolute path within the archive
                        if (target.StartsWith("/"))
                            return target.TrimStart('/');
                        else
                        {
                            var baseDir = Path.GetDirectoryName(worksheetPath)?.Replace("\\", "/") ?? "xl/worksheets";
                            return $"{baseDir}/{target}";
                        }
                    }
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Reads a picture shape from the drawing XML.
    /// </summary>
    private static ShapeInfo? ReadPicture(XmlReader reader, ZipArchive archive, string drawingPath, int shapeId, Action<string>? log)
    {
        try
        {
            var name = "";
            var description = "";
            ShapePositionInfo? position = null;
            ImageInfo? imageInfo = null;

            // Read picture properties
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "pic")
                    break;

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "cNvPr": // Non-visual properties
                            name = reader.GetAttribute("name") ?? $"Picture {shapeId}";
                            description = reader.GetAttribute("descr") ?? "";
                            break;

                        case "blip": // Image reference
                            var embed = reader.GetAttribute("embed");
                            if (!string.IsNullOrEmpty(embed))
                            {
                                imageInfo = ReadImageData(archive, drawingPath, embed, log);
                            }
                            break;

                        case "xfrm": // Transform (position)
                            position = ReadTransform(reader);
                            break;

                        case "from": // Two-cell anchor from
                        case "to": // Two-cell anchor to
                            // Handled by ReadTransform or separate anchor reading
                            break;
                    }
                }
            }

            if (imageInfo == null)
            {
                log?.Invoke($"[ShapeReader] Picture without image data: {name}");
                return null;
            }

            position ??= new ShapePositionInfo(
                AnchorType.OneCell, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

            return new ShapeInfo(
                shapeId,
                ShapeType.Picture,
                name,
                position.Value,
                imageInfo,
                true,
                0,
                shapeId);
        }
        catch (Exception ex)
        {
            log?.Invoke($"[ShapeReader] Error reading picture: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Reads a shape from the drawing XML.
    /// </summary>
    private static ShapeInfo? ReadShape(XmlReader reader, int shapeId, Action<string>? log)
    {
        try
        {
            var name = $"Shape {shapeId}";
            var shapeType = ShapeType.Rectangle;
            ShapePositionInfo? position = null;

            // Read shape properties
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "sp")
                    break;

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "cNvPr":
                            name = reader.GetAttribute("name") ?? name;
                            break;

                        case "prstGeom":
                            var prst = reader.GetAttribute("prst");
                            shapeType = ParseShapeType(prst);
                            break;

                        case "xfrm":
                            position = ReadTransform(reader);
                            break;
                    }
                }
            }

            position ??= new ShapePositionInfo(
                AnchorType.OneCell, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

            return new ShapeInfo(
                shapeId,
                shapeType,
                name,
                position.Value,
                null,
                true,
                0,
                shapeId);
        }
        catch (Exception ex)
        {
            log?.Invoke($"[ShapeReader] Error reading shape: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Reads a connector shape (line/arrow) from the drawing XML.
    /// </summary>
    private static ShapeInfo? ReadConnector(XmlReader reader, int shapeId, Action<string>? log)
    {
        try
        {
            var name = $"Connector {shapeId}";
            ShapePositionInfo? position = null;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "cxnSp")
                    break;

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "cNvPr":
                            name = reader.GetAttribute("name") ?? name;
                            break;

                        case "xfrm":
                            position = ReadTransform(reader);
                            break;
                    }
                }
            }

            position ??= new ShapePositionInfo(
                AnchorType.OneCell, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0);

            return new ShapeInfo(
                shapeId,
                ShapeType.Line,
                name,
                position.Value,
                null,
                true,
                0,
                shapeId);
        }
        catch (Exception ex)
        {
            log?.Invoke($"[ShapeReader] Error reading connector: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Reads image data from the image part.
    /// </summary>
    private static ImageInfo? ReadImageData(ZipArchive archive, string drawingPath, string relationshipId, Action<string>? log)
    {
        try
        {
            // Get the image path from drawing relationships
            var drawingDir = Path.GetDirectoryName(drawingPath)?.Replace("\\", "/") ?? "xl/drawings";
            var relsPath = $"{drawingDir}/_rels/{Path.GetFileName(drawingPath)}.rels";
            var relsEntry = archive.GetEntry(relsPath);

            if (relsEntry == null)
            {
                log?.Invoke($"[ShapeReader] Drawing relationships not found: {relsPath}");
                return null;
            }

            string? imagePath = null;
            string? contentType = null;

            using (var stream = relsEntry.Open())
            using (var reader = XmlReader.Create(stream))
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship")
                    {
                        var id = reader.GetAttribute("Id");
                        if (id == relationshipId)
                        {
                            var target = reader.GetAttribute("Target");
                            var type = reader.GetAttribute("Type");

                            if (!string.IsNullOrEmpty(target))
                            {
                                if (target.StartsWith("/"))
                                    imagePath = target.TrimStart('/');
                                else
                                    imagePath = $"{drawingDir}/{target}";
                            }

                            if (type?.Contains("image") == true)
                            {
                                contentType = GetContentTypeFromRelationshipType(type);
                            }

                            break;
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(imagePath))
            {
                log?.Invoke($"[ShapeReader] Image path not found for relationship: {relationshipId}");
                return null;
            }

            // Read image data
            var imageEntry = archive.GetEntry(imagePath);
            if (imageEntry == null)
            {
                log?.Invoke($"[ShapeReader] Image entry not found: {imagePath}");
                return null;
            }

            byte[] imageData;
            using (var stream = imageEntry.Open())
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                imageData = ms.ToArray();
            }

            var format = DetectImageFormat(imageData, imagePath);
            contentType ??= GetContentTypeFromFormat(format);

            return new ImageInfo(format, imageData, relationshipId, contentType);
        }
        catch (Exception ex)
        {
            log?.Invoke($"[ShapeReader] Error reading image data: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Reads transform/position information from the drawing XML.
    /// </summary>
    private static ShapePositionInfo ReadTransform(XmlReader reader)
    {
        var anchorType = AnchorType.OneCell;
        var fromCol = 0;
        var fromRow = 0;
        var fromColOffset = 0;
        var fromRowOffset = 0;
        var toCol = 0;
        var toRow = 0;
        var toColOffset = 0;
        var toRowOffset = 0;
        var widthEmu = 0;
        var heightEmu = 0;

        // Check parent element to determine anchor type
        if (reader.LocalName == "xfrm")
        {
            // Read extents
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "xfrm")
                    break;

                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.LocalName)
                    {
                        case "ext":
                            var cx = reader.GetAttribute("cx");
                            var cy = reader.GetAttribute("cy");
                            if (int.TryParse(cx, out var w)) widthEmu = w;
                            if (int.TryParse(cy, out var h)) heightEmu = h;
                            break;

                        case "off":
                            var x = reader.GetAttribute("x");
                            var y = reader.GetAttribute("y");
                            if (int.TryParse(x, out var xEmu)) fromColOffset = xEmu;
                            if (int.TryParse(y, out var yEmu)) fromRowOffset = yEmu;
                            break;
                    }
                }
            }
        }

        return new ShapePositionInfo(
            anchorType, fromCol, fromRow, fromColOffset, fromRowOffset,
            toCol, toRow, toColOffset, toRowOffset, widthEmu, heightEmu, 0, 0);
    }

    /// <summary>
    /// Parses shape type from preset geometry name.
    /// </summary>
    private static ShapeType ParseShapeType(string? preset)
    {
        return preset?.ToLowerInvariant() switch
        {
            "rect" or "rectangle" => ShapeType.Rectangle,
            "ellipse" or "oval" => ShapeType.Oval,
            "arc" => ShapeType.Arc,
            "line" => ShapeType.Line,
            "straightConnector1" or "connector" => ShapeType.Line,
            "arrow" => ShapeType.Arrow,
            "textBox" => ShapeType.TextBox,
            _ => ShapeType.Rectangle
        };
    }

    /// <summary>
    /// Detects image format from file header or extension.
    /// </summary>
    private static ImageFormat DetectImageFormat(byte[] data, string path)
    {
        if (data.Length >= 8)
        {
            // PNG signature
            if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
                return ImageFormat.Png;

            // JPEG signature
            if (data[0] == 0xFF && data[1] == 0xD8)
                return ImageFormat.Jpeg;

            // GIF signature
            if (data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46)
                return ImageFormat.Gif;

            // BMP signature
            if (data[0] == 0x42 && data[1] == 0x4D)
                return ImageFormat.Bmp;

            // TIFF signature
            if ((data[0] == 0x49 && data[1] == 0x49 && data[2] == 0x2A && data[3] == 0x00) ||
                (data[0] == 0x4D && data[1] == 0x4D && data[2] == 0x00 && data[3] == 0x2A))
                return ImageFormat.Tiff;
        }

        // Fallback to extension
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".png" => ImageFormat.Png,
            ".jpg" or ".jpeg" => ImageFormat.Jpeg,
            ".gif" => ImageFormat.Gif,
            ".bmp" => ImageFormat.Bmp,
            ".tif" or ".tiff" => ImageFormat.Tiff,
            ".wmf" => ImageFormat.Wmf,
            ".emf" => ImageFormat.Emf,
            _ => ImageFormat.Unknown
        };
    }

    /// <summary>
    /// Gets content type from relationship type.
    /// </summary>
    private static string? GetContentTypeFromRelationshipType(string type)
    {
        if (type.Contains("png")) return "image/png";
        if (type.Contains("jpeg") || type.Contains("jpg")) return "image/jpeg";
        if (type.Contains("gif")) return "image/gif";
        if (type.Contains("bmp")) return "image/bmp";
        if (type.Contains("tiff")) return "image/tiff";
        return null;
    }

    /// <summary>
    /// Gets content type from image format.
    /// </summary>
    private static string? GetContentTypeFromFormat(ImageFormat format)
    {
        return format switch
        {
            ImageFormat.Png => "image/png",
            ImageFormat.Jpeg => "image/jpeg",
            ImageFormat.Gif => "image/gif",
            ImageFormat.Bmp => "image/bmp",
            ImageFormat.Tiff => "image/tiff",
            ImageFormat.Wmf => "image/x-wmf",
            ImageFormat.Emf => "image/x-emf",
            _ => null
        };
    }
}
