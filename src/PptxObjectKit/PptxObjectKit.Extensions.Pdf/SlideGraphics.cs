using System.Xml.Linq;
using SkiaSharp;
using static PptxObjectKit.Graphs.PptxPackageGraph;

namespace PptxObjectKit.Extensions.Pdf;

internal sealed partial class SlideRenderer
{
    private const float Emu = 12700;

    internal void Render(SlideScene slide, SKCanvas canvas, float width, float height)
    {
        scene = slide;
        canvas.Clear(SKColors.White);
        canvas.ClipRect(new SKRect(0, 0, width, height));
        var background = D(slide.Xml, "bgPr").FirstOrDefault() ?? D(slide.Layout, "bgPr").FirstOrDefault() ?? D(slide.Master, "bgPr").FirstOrDefault();
        using (var paint = Fill(background, width, height))
        {
            if (paint is not null)
            {
                canvas.DrawRect(0, 0, width, height, paint);
            }
        }
        if (D(slide.Xml, "timing").Any())
        {
            Warn("Animations are exported as a static slide.");
        }
        if (D(slide.Xml, "transition").Any())
        {
            Warn("Slide transitions are not included in PDF.");
        }
        if (D(slide.Xml, "bgRef").Any())
        {
            Warn("Theme background references are not yet supported.");
        }
        if (N(slide.Xml.Root, "showMasterSp", 1) != 0)
        {
            foreach (var (xml, part) in new[] { (slide.Master, slide.MasterPart), (slide.Layout, slide.LayoutPart) })
            {
                foreach (var item in D(xml, "spTree").FirstOrDefault()?.Elements() ?? [])
                {
                    if (!D(item, "ph").Any())
                    {
                        Draw(canvas, item, part!, 0, false);
                    }
                }
            }
        }
        foreach (var item in D(slide.Xml, "spTree").First().Elements())
        {
            Draw(canvas, item, slide.Part, 0, true);
        }
    }

    private SKColor Color(XElement? node, SKColor? placeholder = null, int depth = 0)
    {
        if (depth > 16)
        {
            throw new InvalidDataException("Cyclic or excessively nested theme colors.");
        }
        if (node is null)
        {
            return SKColors.Black;
        }
        var value = (string?)node.Attribute("val") ?? "000000";
        SKColor color;
        if (node.Name.LocalName == "schemeClr")
        {
            if (value == "phClr")
            {
                return placeholder ?? SKColors.Black;
            }
            var map = D(scene.Xml, "overrideClrMapping").FirstOrDefault() ?? D(scene.Master, "clrMap").FirstOrDefault();
            value = (string?)map?.Attribute(value) ?? value switch { "tx1" => "dk1", "tx2" => "dk2", "bg1" => "lt1", "bg2" => "lt2", _ => value };
            color = Color(C(D(scene.Theme, "clrScheme").FirstOrDefault(), value)?.Elements().FirstOrDefault(), depth: depth + 1);
        }
        else if (node.Name.LocalName == "sysClr")
        {
            color = SKColor.TryParse("#" + ((string?)node.Attribute("lastClr") ?? "000000"), out var parsed) ? parsed : SKColors.Black;
        }
        else if (node.Name.LocalName == "prstClr")
        {
            color = value switch { "white" => SKColors.White, "red" => SKColors.Red, "blue" => SKColors.Blue, "green" => SKColors.Green, "yellow" => SKColors.Yellow, "gray" => SKColors.Gray, _ => SKColors.Black };
        }
        else
        {
            color = SKColor.TryParse("#" + value, out var parsed) ? parsed : SKColors.Black;
        }
        float r = color.Red, g = color.Green, b = color.Blue, alpha = color.Alpha;
        foreach (var modifier in node.Elements())
        {
            var amount = N(modifier, "val") / 100000;
            switch (modifier.Name.LocalName)
            {
                case "alpha": alpha = amount * 255; break;
                case "alphaMod": alpha *= amount; break;
                case "shade": case "lumMod": r *= amount; g *= amount; b *= amount; break;
                case "tint": r += (255 - r) * amount; g += (255 - g) * amount; b += (255 - b) * amount; break;
                case "lumOff": r += amount * 255; g += amount * 255; b += amount * 255; break;
            }
        }
        return new((byte)Math.Clamp(r, 0, 255), (byte)Math.Clamp(g, 0, 255), (byte)Math.Clamp(b, 0, 255), (byte)Math.Clamp(alpha, 0, 255));
    }

    private SKPaint? Fill(XElement? properties, float width, float height, SKColor? placeholder = null)
    {
        if (properties is null || C(properties, "noFill") is not null)
        {
            return null;
        }
        if (C(properties, "solidFill") is { } solid)
        {
            return new SKPaint { Color = Color(solid.Elements().FirstOrDefault(), placeholder), IsAntialias = true };
        }
        if (C(properties, "gradFill") is { } gradient)
        {
            var stops = D(gradient, "gs").OrderBy(item => N(item, "pos")).ToArray();
            if (stops.Length == 0)
            {
                return null;
            }
            var colors = stops.Select(item => Color(item.Elements().FirstOrDefault(), placeholder)).ToArray();
            var positions = stops.Select(item => Math.Clamp(N(item, "pos") / 100000, 0, 1)).ToArray();
            var angle = N(C(gradient, "lin"), "ang") / 60000 * MathF.PI / 180;
            var dx = MathF.Cos(angle); var dy = MathF.Sin(angle);
            var length = MathF.Abs(width * dx) + MathF.Abs(height * dy);
            var shader = C(gradient, "path") is not null
                ? SKShader.CreateRadialGradient(new(width / 2, height / 2), Math.Max(1, Math.Max(width, height) / 2), colors, positions, SKShaderTileMode.Clamp)
                : SKShader.CreateLinearGradient(new(width / 2 - dx * length / 2, height / 2 - dy * length / 2), new(width / 2 + dx * length / 2, height / 2 + dy * length / 2), colors, positions, SKShaderTileMode.Clamp);
            return new SKPaint { Shader = shader, IsAntialias = true };
        }
        if (C(properties, "pattFill") is not null || C(properties, "blipFill") is not null)
        {
            Warn("Pattern/picture shape fills are not yet supported.");
        }
        return null;
    }

    private void Draw(SKCanvas canvas, XElement original, string part, int depth, bool resolve)
    {
        cancellation.ThrowIfCancellationRequested();
        if (depth > 64)
        {
            throw new NotSupportedException("Shape nesting exceeds 64 levels.");
        }
        var kind = original.Name.LocalName;
        if (kind is "nvGrpSpPr" or "grpSpPr" or "extLst")
        {
            return;
        }
        if (N(D(original, "cNvPr").FirstOrDefault(), "hidden") != 0)
        {
            return;
        }
        var shape = resolve && kind == "sp" ? Prepare(original) : new XElement(original);
        var props = C(shape, kind == "grpSp" ? "grpSpPr" : "spPr");
        var frame = kind == "graphicFrame" ? C(shape, "xfrm") : C(props, "xfrm");
        if (frame is null)
        {
            Warn($"{kind} without a resolved transform was omitted.");
            return;
        }
        var width = N(C(frame, "ext"), "cx") / Emu;
        var height = N(C(frame, "ext"), "cy") / Emu;
        var x = N(C(frame, "off"), "x") / Emu;
        var y = N(C(frame, "off"), "y") / Emu;
        canvas.Save();
        try
        {
            canvas.Translate(x + width / 2, y + height / 2);
            canvas.RotateDegrees(N(frame, "rot") / 60000);
            canvas.Scale(N(frame, "flipH") != 0 ? -1 : 1, N(frame, "flipV") != 0 ? -1 : 1);
            canvas.Translate(-width / 2, -height / 2);
            if (D(shape, "effectLst").Any(item => item.HasElements) || D(shape, "scene3d").Any())
            {
                Warn("Shadows, 3D and advanced effects are not reproduced.");
            }
            if (kind == "grpSp")
            {
                var cx = N(C(frame, "chExt"), "cx") / Emu;
                var cy = N(C(frame, "chExt"), "cy") / Emu;
                if (cx == 0 || cy == 0)
                {
                    Warn("A group with zero child extent was omitted.");
                    return;
                }
                canvas.Scale(width / cx, height / cy);
                canvas.Translate(-N(C(frame, "chOff"), "x") / Emu, -N(C(frame, "chOff"), "y") / Emu);
                foreach (var child in shape.Elements())
                {
                    if (C(C(child, "spPr"), "grpFill") is { } groupFill)
                    {
                        var fill = props?.Elements().FirstOrDefault(item => item.Name.LocalName is "solidFill" or "gradFill" or "noFill");
                        if (fill is not null)
                        {
                            groupFill.ReplaceWith(new XElement(fill));
                        }
                    }
                    Draw(canvas, child, part, depth + 1, resolve);
                }
                return;
            }
            if (kind == "pic")
            {
                Picture(canvas, shape, part, width, height);
                return;
            }
            if (kind == "graphicFrame")
            {
                if (D(shape, "tbl").FirstOrDefault() is { } table)
                {
                    Table(canvas, table, width, height);
                }
                else if (D(shape, "relIds").FirstOrDefault() is { } ids)
                {
                    var dataPart = Related(part, "diagramData", (string?)ids.Attribute(R + "dm"));
                    var model = dataPart is null ? null : document.Xml(dataPart);
                    var drawingId = (string?)D(model, "dataModelExt").FirstOrDefault()?.Attribute("relId");
                    var drawingPart = drawingId is null ? null : Related(part, "diagramDrawing", drawingId) ?? Related(dataPart!, "diagramDrawing", drawingId);
                    var tree = drawingPart is null ? null : D(document.Xml(drawingPart), "spTree").FirstOrDefault();
                    if (tree is null)
                    {
                        Warn("SmartArt without a saved drawing cache is not supported.");
                    }
                    else
                    {
                        foreach (var child in tree.Elements())
                        {
                            Draw(canvas, child, drawingPart!, depth + 1, true);
                        }
                    }
                }
                else
                {
                    Warn("Charts and unsupported graphic frames are not yet supported.");
                }
                return;
            }
            if (kind is not ("sp" or "cxnSp"))
            {
                Warn($"Unsupported shape {kind} was omitted.");
                return;
            }
            using var path = ShapePath(shape, width, height);
            using var fillPaint = ShapePaint(shape, false, width, height);
            if (fillPaint is not null && kind != "cxnSp")
            {
                canvas.DrawPath(path, fillPaint);
            }
            using var stroke = ShapePaint(shape, true, width, height);
            if (stroke is not null)
            {
                stroke.Style = SKPaintStyle.Stroke;
                var lineIndex = (int)N(C(C(shape, "style"), "lnRef"), "idx") - 1;
                var themeLine = D(scene.Theme, "lnStyleLst").FirstOrDefault()?.Elements().ElementAtOrDefault(lineIndex);
                stroke.StrokeWidth = Math.Max(.1f, N(C(props, "ln"), "w", N(themeLine, "w", Emu)) / Emu);
                canvas.DrawPath(path, stroke);
            }
            if (stroke is not null)
            {
                using var measure = new SKPathMeasure(path, false);
                if (measure.GetPositionAndTangent(0, out var start, out var startTangent))
                {
                    Arrow(canvas, C(C(props, "ln"), "headEnd"), start.X, start.Y, start.X + startTangent.X, start.Y + startTangent.Y, stroke);
                }
                if (measure.GetPositionAndTangent(measure.Length, out var end, out var endTangent))
                {
                    Arrow(canvas, C(C(props, "ln"), "tailEnd"), end.X, end.Y, end.X - endTangent.X, end.Y - endTangent.Y, stroke);
                }
            }
            if (C(shape, "txBody") is { } body)
            {
                // DrawingML flips shape geometry without mirroring its text.
                canvas.Translate(width / 2, height / 2);
                canvas.Scale(N(frame, "flipH") != 0 ? -1 : 1, N(frame, "flipV") != 0 ? -1 : 1);
                canvas.Translate(-width / 2, -height / 2);
                Text(canvas, body, width, height);
            }
        }
        finally
        {
            canvas.Restore();
        }
    }

    private SKPaint? ShapePaint(XElement shape, bool line, float width, float height)
    {
        var props = C(shape, "spPr");
        var direct = line ? C(props, "ln") : props;
        if (direct?.Elements().Any(item => item.Name.LocalName.EndsWith("Fill", StringComparison.Ordinal)) == true)
        {
            return Fill(direct, width, height);
        }
        var reference = C(C(shape, "style"), line ? "lnRef" : "fillRef");
        var index = (int)N(reference, "idx");
        var name = line ? "lnStyleLst" : index >= 1001 ? "bgFillStyleLst" : "fillStyleLst";
        var item = D(scene.Theme, name).FirstOrDefault()?.Elements().ElementAtOrDefault(index >= 1001 ? index - 1001 : index - 1);
        if (item is null)
        {
            return null;
        }
        return Fill(line ? item : new XElement("properties", new XElement(item)), width, height, Color(reference?.Elements().FirstOrDefault()));
    }

    private void Arrow(SKCanvas canvas, XElement? end, float x, float y, float towardX, float towardY, SKPaint stroke)
    {
        var type = (string?)end?.Attribute("type") ?? "none";
        if (type == "none")
        {
            return;
        }
        if (type is not ("triangle" or "arrow" or "diamond" or "oval" or "stealth"))
        {
            Warn($"Arrowhead '{type}' is not supported.");
            return;
        }
        var length = Math.Max(3, stroke.StrokeWidth * ((string?)end?.Attribute("len") == "lg" ? 5 : 3));
        var halfWidth = Math.Max(2, stroke.StrokeWidth * ((string?)end?.Attribute("w") == "lg" ? 2.5f : 1.5f));
        canvas.Save();
        canvas.Translate(x, y);
        canvas.RotateDegrees(MathF.Atan2(towardY - y, towardX - x) * 180 / MathF.PI);
        using var path = new SKPath();
        if (type == "oval")
        {
            path.AddOval(new(-length / 2, -halfWidth, length / 2, halfWidth));
        }
        else if (type == "diamond")
        {
            path.MoveTo(0, 0);
            path.LineTo(length / 2, -halfWidth);
            path.LineTo(length, 0);
            path.LineTo(length / 2, halfWidth);
            path.Close();
        }
        else
        {
            path.MoveTo(length, -halfWidth);
            path.LineTo(0, 0);
            path.LineTo(length, halfWidth);
            if (type != "arrow")
            {
                path.Close();
            }
        }
        using var paint = new SKPaint { Color = stroke.Color, Shader = stroke.Shader, IsAntialias = true, Style = type == "arrow" ? SKPaintStyle.Stroke : SKPaintStyle.Fill, StrokeWidth = stroke.StrokeWidth };
        canvas.DrawPath(path, paint);
        canvas.Restore();
    }

    private SKPath ShapePath(XElement shape, float width, float height)
    {
        var props = C(shape, "spPr");
        var path = new SKPath();
        if (C(props, "custGeom") is { } custom)
        {
            foreach (var source in D(custom, "path"))
            {
                var sx = width / Math.Max(1, N(source, "w", 1));
                var sy = height / Math.Max(1, N(source, "h", 1));
                foreach (var segment in source.Elements())
                {
                    var pts = D(segment, "pt").Select(point => new SKPoint(N(point, "x") * sx, N(point, "y") * sy)).ToArray();
                    switch (segment.Name.LocalName)
                    {
                        case "moveTo" when pts.Length == 1: path.MoveTo(pts[0]); break;
                        case "lnTo" when pts.Length == 1: path.LineTo(pts[0]); break;
                        case "cubicBezTo" when pts.Length == 3: path.CubicTo(pts[0], pts[1], pts[2]); break;
                        case "quadBezTo" when pts.Length == 2: path.QuadTo(pts[0], pts[1]); break;
                        case "close": path.Close(); break;
                        default: Warn("Unsupported custom geometry segment was omitted."); break;
                    }
                }
            }
            return path;
        }
        var preset = (string?)C(props, "prstGeom")?.Attribute("prst") ?? "rect";
        switch (preset)
        {
            case "ellipse": path.AddOval(new(0, 0, width, height)); break;
            case "round2DiagRect":
                var radius = Math.Min(width, height) * .16667f;
                path.MoveTo(radius, 0);
                path.LineTo(width, 0);
                path.LineTo(width, height - radius);
                path.QuadTo(width, height, width - radius, height);
                path.LineTo(0, height);
                path.LineTo(0, radius);
                path.QuadTo(0, 0, radius, 0);
                path.Close();
                break;
            case "leftBrace":
                path.MoveTo(width, 0);
                path.CubicTo(0, 0, width, height / 2, 0, height / 2);
                path.CubicTo(width, height / 2, 0, height, width, height);
                break;
            case "roundRect": path.AddRoundRect(new(0, 0, width, height), Math.Min(width, height) * .17f, Math.Min(width, height) * .17f); break;
            case "line": case "straightConnector1": path.MoveTo(0, 0); path.LineTo(width, height); break;
            case "bentConnector2": path.MoveTo(0, 0); path.LineTo(width, 0); path.LineTo(width, height); break;
            case "bentConnector3": path.MoveTo(0, 0); path.LineTo(width / 2, 0); path.LineTo(width / 2, height); path.LineTo(width, height); break;
            case "triangle": path.MoveTo(width / 2, 0); path.LineTo(width, height); path.LineTo(0, height); path.Close(); break;
            case "diamond": path.MoveTo(width / 2, 0); path.LineTo(width, height / 2); path.LineTo(width / 2, height); path.LineTo(0, height / 2); path.Close(); break;
            default:
                path.AddRect(new(0, 0, width, height));
                if (preset != "rect")
                {
                    Warn($"Preset '{preset}' uses a rectangular fallback.");
                }
                break;
        }
        return path;
    }
}
