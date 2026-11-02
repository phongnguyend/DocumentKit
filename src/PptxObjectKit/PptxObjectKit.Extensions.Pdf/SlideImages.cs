using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using SkiaSharp;
using static PptxObjectKit.Graphs.PptxPackageGraph;

namespace PptxObjectKit.Extensions.Pdf;

internal sealed partial class SlideRenderer
{
    private void Picture(SKCanvas canvas, XElement shape, string part, float width, float height)
    {
        var blip = D(shape, "blip").FirstOrDefault();
        var svg = D(shape, "svgBlip").FirstOrDefault();
        var id = (string?)blip?.Attribute(R + "embed") ?? (string?)svg?.Attribute(R + "embed");
        var target = id is null ? null : Related(part, "image", id);
        if (target is null)
        {
            Warn("A missing or externally linked picture was omitted; external URLs are never fetched.");
            return;
        }
        var bytes = document.GetPart(target);
        if (target.EndsWith(".svg", StringComparison.OrdinalIgnoreCase))
        {
            Svg(canvas, bytes, width, height);
            return;
        }
        using var data = SKData.CreateCopy(bytes);
        using var codec = SKCodec.Create(data);
        if (codec is null)
        {
            Warn($"Picture format '{Path.GetExtension(target)}' is not supported.");
            return;
        }
        if ((long)codec.Info.Width * codec.Info.Height > 64000000)
        {
            throw new NotSupportedException("A picture exceeds the 64 megapixel decoding limit.");
        }
        using var image = SKImage.FromEncodedData(data);
        if (image is null)
        {
            Warn("A picture could not be decoded.");
            return;
        }
        var crop = D(shape, "srcRect").FirstOrDefault();
        var rect = new SKRect(image.Width * N(crop, "l") / 100000, image.Height * N(crop, "t") / 100000, image.Width * (1 - N(crop, "r") / 100000), image.Height * (1 - N(crop, "b") / 100000));
        using var paint = new SKPaint { Color = SKColors.White.WithAlpha((byte)Math.Clamp(N(C(blip, "alphaModFix"), "amt", 100000) / 100000 * 255, 0, 255)), IsAntialias = true };
        canvas.DrawImage(image, rect, new SKRect(0, 0, width, height), new SKSamplingOptions(SKFilterMode.Linear), paint);
    }

    private void Svg(SKCanvas canvas, byte[] bytes, float width, float height)
    {
        using var stream = new MemoryStream(bytes);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 4 * 1024 * 1024 });
        var xml = XDocument.Load(reader);
        var root = xml.Root!;
        var bounds = Numbers((string?)root.Attribute("viewBox") ?? "");
        var sw = bounds.Length == 4 ? bounds[2] : N(root, "width", width);
        var sh = bounds.Length == 4 ? bounds[3] : N(root, "height", height);
        if (sw <= 0 || sh <= 0)
        {
            Warn("An SVG with invalid dimensions was omitted.");
            return;
        }
        canvas.Save();
        canvas.ClipRect(new(0, 0, width, height));
        canvas.Scale(width / sw, height / sh);
        if (bounds.Length == 4)
        {
            canvas.Translate(-bounds[0], -bounds[1]);
        }
        Visit(root, SKColors.Black, 0);
        canvas.Restore();

        void Visit(XElement element, SKColor inherited, int depth)
        {
            if (depth > 64)
            {
                throw new NotSupportedException("SVG nesting exceeds 64 levels.");
            }
            var fill = (string?)element.Attribute("fill");
            var color = fill == "none" ? SKColors.Transparent : fill is null ? inherited : SKColor.TryParse(fill, out var parsed) ? parsed : inherited;
            color = color.WithAlpha((byte)Math.Clamp(color.Alpha * N(element, "opacity", 1) * N(element, "fill-opacity", 1), 0, 255));
            canvas.Save();
            var transform = (string?)element.Attribute("transform") ?? "";
            foreach (Match match in Regex.Matches(transform, @"(matrix|translate|scale|rotate)\s*\(([^)]*)\)"))
            {
                var values = Numbers(match.Groups[2].Value);
                switch (match.Groups[1].Value)
                {
                    case "translate" when values.Length > 0: canvas.Translate(values[0], values.Length > 1 ? values[1] : 0); break;
                    case "scale" when values.Length > 0: canvas.Scale(values[0], values.Length > 1 ? values[1] : values[0]); break;
                    case "rotate" when values.Length > 0: canvas.RotateDegrees(values[0], values.Length > 2 ? values[1] : 0, values.Length > 2 ? values[2] : 0); break;
                    case "matrix" when values.Length == 6:
                        var matrix = new SKMatrix(values[0], values[2], values[4], values[1], values[3], values[5], 0, 0, 1);
                        canvas.Concat(in matrix);
                        break;
                }
            }
            using var paint = new SKPaint { Color = color, IsAntialias = true };
            switch (element.Name.LocalName)
            {
                case "svg": case "g":
                    foreach (var child in element.Elements())
                    {
                        Visit(child, color, depth + 1);
                    }
                    break;
                case "path":
                    using (var path = SKPath.ParseSvgPathData((string?)element.Attribute("d") ?? ""))
                    {
                        if (path is not null)
                        {
                            path.FillType = (string?)element.Attribute("fill-rule") == "evenodd" ? SKPathFillType.EvenOdd : SKPathFillType.Winding;
                            canvas.DrawPath(path, paint);
                            if ((string?)element.Attribute("stroke") is { } stroke && SKColor.TryParse(stroke, out var strokeColor))
                            {
                                paint.Color = strokeColor;
                                paint.Style = SKPaintStyle.Stroke;
                                paint.StrokeWidth = N(element, "stroke-width", 1);
                                canvas.DrawPath(path, paint);
                            }
                        }
                    }
                    break;
                case "rect": canvas.DrawRect(N(element, "x"), N(element, "y"), N(element, "width"), N(element, "height"), paint); break;
                case "circle": canvas.DrawCircle(N(element, "cx"), N(element, "cy"), N(element, "r"), paint); break;
                case "ellipse": canvas.DrawOval(N(element, "cx"), N(element, "cy"), N(element, "rx"), N(element, "ry"), paint); break;
                case "title": case "desc": break;
                default: Warn($"SVG element '{element.Name.LocalName}' is not supported."); break;
            }
            canvas.Restore();
        }
    }

    private static float[] Numbers(string value)
    {
        var numbers = Regex.Matches(value, @"[-+]?(?:\d*\.)?\d+(?:[eE][-+]?\d+)?").Select(match => float.Parse(match.Value, CultureInfo.InvariantCulture)).ToArray();
        if (numbers.Any(number => !float.IsFinite(number) || Math.Abs(number) > 1e12))
        {
            throw new InvalidDataException("Invalid SVG numeric value.");
        }
        return numbers;
    }
}
