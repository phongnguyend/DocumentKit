using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using PdfObjectKit.Interpreters;
using PdfObjectKit.Models;
using PdfObjectKit.Options;
using PdfObjectKit.Rendering;
using static PdfObjectKit.Converters.PptxPackage;

namespace PdfObjectKit.Converters;

internal sealed class PptxSlideWriter(PptxPackage package, PdfMatrix pageTransform, double slideWidth, double slideHeight, PdfPptxOptions options, Action<string> warn)
{
    private sealed record State(PdfMatrix Matrix, string Fill = "000000", string Stroke = "000000", double FillAlpha = 1, double StrokeAlpha = 1, double LineWidth = 1, int TextMode = 0, PptxGradient? FillGradient = null, PptxGradient? StrokeGradient = null, int ClipBoundary = 0);
    private readonly XElement tree = Tree();
    private readonly List<XElement> relationships = [];
    private readonly List<(string Kind, PdfImagePoint[] Points)> path = [];
    private readonly Stack<State> stack = new();
    private State state = new(PdfMatrix.Identity);
    private PdfImagePoint current = new(0, 0);
    private PdfImagePoint subpathStart = new(0, 0);
    private int objectId = 1;
    private int clipBoundary;
    private sealed record Outline(XElement Shape, State PaintState, double Left, double Top, double Width, double Height);
    private Outline? lastOutline;
    private readonly Dictionary<XElement, (PdfCanvasCommand Command, State PaintState)> outlinedText = [];
    private XElement? textInsertionAnchor;
    private readonly List<TextFragment> textLine = [];
    private static readonly PdfHeuristicOptions TextHeuristics = new();
    private sealed record TextFragment(string Text, XElement Properties, XElement Frame, PdfImagePoint Start, PdfImagePoint End, PdfImagePoint Direction, PdfImagePoint Up, double Height, PptxGradient? Gradient, double Alpha);

    internal int TextCount { get; private set; }

    internal int ImageCount { get; private set; }

    internal int ShapeCount { get; private set; }

    internal void Write(IReadOnlyList<PdfCanvasCommand> commands, CancellationToken cancellation)
    {
        foreach (var command in commands)
        {
            cancellation.ThrowIfCancellationRequested();
            if (objectId > 50000 || textLine.Count > 50000 || path.Count > 200000)
            {
                throw new NotSupportedException("PowerPoint slide object/path limit exceeded.");
            }
            var n = command.Numbers;
            if (n.Any(value => !double.IsFinite(value)))
            {
                throw new InvalidDataException("Non-finite drawing coordinates.");
            }
            var matrix = pageTransform.Multiply(state.Matrix);
            PdfImagePoint Point(double x, double y) => matrix.Point(x, y);
            // State saves/restores do not paint anything and must not split text lines.
            if (command.Kind is "jpeg" or "rgba" or "s" or "b" or "b*" or "S" or "f" or "F" or "f*" or "B" or "B*")
            {
                FlushText();
            }
            switch (command.Kind)
            {
                case "q":
                    stack.Push(state);
                    break;
                case "Q":
                    var restored = stack.Count > 0 ? stack.Pop() : new(PdfMatrix.Identity);
                    if (state.ClipBoundary != restored.ClipBoundary)
                    {
                        FlushText();
                    }
                    state = restored;
                    break;
                case "cm":
                    state = state with { Matrix = state.Matrix.Multiply(new(n[0], n[1], n[2], n[3], n[4], n[5])) };
                    break;
                case "rg": case "g": case "k":
                    state = state with { Fill = Color(n), FillGradient = null };
                    break;
                case "RG": case "G": case "K":
                    state = state with { Stroke = Color(n), StrokeGradient = null };
                    break;
                case "alpha":
                    state = state with { FillAlpha = n[0] < 0 ? state.FillAlpha : Math.Clamp(n[0], 0, 1), StrokeAlpha = n[1] < 0 ? state.StrokeAlpha : Math.Clamp(n[1], 0, 1) };
                    break;
                case "Tr":
                    state = state with { TextMode = (int)n[0] };
                    if (n[0] >= 4)
                    {
                        warn("Text clipping is not preserved.");
                    }
                    break;
                case "w":
                    state = state with { LineWidth = n[0] };
                    break;
                case "text":
                    Text(command);
                    break;
                case "jpeg": case "rgba":
                    if (options.IncludeImages)
                    {
                        Picture(command, matrix);
                    }
                    break;
                case "m": case "l":
                    current = Point(n[0], n[1]);
                    if (command.Kind == "m")
                    {
                        subpathStart = current;
                    }
                    path.Add((command.Kind == "m" ? "moveTo" : "lnTo", [current]));
                    break;
                case "c":
                    current = Point(n[4], n[5]);
                    path.Add(("cubicBezTo", [Point(n[0], n[1]), Point(n[2], n[3]), current]));
                    break;
                case "v":
                    path.Add(("cubicBezTo", [current, Point(n[0], n[1]), Point(n[2], n[3])]));
                    current = Point(n[2], n[3]);
                    break;
                case "y":
                    current = Point(n[2], n[3]);
                    path.Add(("cubicBezTo", [Point(n[0], n[1]), current, current]));
                    break;
                case "re":
                    path.Add(("moveTo", [Point(n[0], n[1])]));
                    path.Add(("lnTo", [Point(n[0] + n[2], n[1])]));
                    path.Add(("lnTo", [Point(n[0] + n[2], n[1] + n[3])]));
                    path.Add(("lnTo", [Point(n[0], n[1] + n[3])]));
                    path.Add(("close", []));
                    current = Point(n[0], n[1]);
                    subpathStart = current;
                    break;
                case "h":
                    current = subpathStart;
                    path.Add(("close", []));
                    break;
                case "s": case "b": case "b*":
                    path.Add(("close", []));
                    Paint(command.Kind, matrix);
                    break;
                case "S": case "f": case "F": case "f*": case "B": case "B*": case "n":
                    Paint(command.Kind, matrix);
                    break;
                case "W": case "W*": case "formClip":
                    if (command.Kind == "formClip" || !IsFullSlideClip())
                    {
                        FlushText();
                        state = state with { ClipBoundary = ++clipBoundary };
                    }
                    warn("Clipping paths are not preserved; objects outside PDF clips may become visible. Review the exported slides.");
                    break;
                case "fillGradient": case "strokeGradient":
                    var gradient = new PptxGradient(n, pageTransform);
                    state = command.Kind == "fillGradient" ? state with { FillGradient = gradient } : state with { StrokeGradient = gradient };
                    break;
                case "d":
                    if (n.Length > 1)
                    {
                        warn("Dashed strokes are exported as solid lines.");
                    }
                    break;
                case "J": case "j": case "M":
                    break;
                default:
                    warn($"Drawing command '{command.Kind}' is not exported.");
                    break;
            }
        }
        FlushText();
        RecoverOutlinedText();
        package.AddSlide(tree, relationships);
    }

    private bool IsFullSlideClip()
    {
        // Repeated rectangular page clips are common around individual words.
        // Only an axis-aligned rectangle covering the entire slide is a no-op.
        if (path.Count != 5 || path[0].Kind != "moveTo" || path[4].Kind != "close"
            || path.Skip(1).Take(3).Any(segment => segment.Kind != "lnTo"))
        {
            return false;
        }
        var points = path.Take(4).Select(segment => segment.Points[0]).ToArray();
        for (var i = 0; i < 4; i++)
        {
            var delta = Difference(points[i], points[(i + 1) % 4]);
            if (Math.Abs(delta.X) > .01 && Math.Abs(delta.Y) > .01)
            {
                return false;
            }
        }
        return points.Min(point => point.X) <= .01 && points.Min(point => point.Y) <= .01
            && points.Max(point => point.X) >= slideWidth - .01 && points.Max(point => point.Y) >= slideHeight - .01;
    }

    private void Text(PdfCanvasCommand command)
    {
        var alpha = state.TextMode is 1 or 5 ? state.StrokeAlpha : state.FillAlpha;
        if (state.TextMode is 3 or 7 || alpha <= 0)
        {
            if (alpha <= 0 && state.TextMode == 0)
            {
                MatchOutlinedText(command);
            }
            return;
        }
        lastOutline = null;
        var text = Clean(string.Concat(command.Glyphs?.Select(glyph => glyph.Text) ?? []));
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }
        // PDF text matrices already include the graphics transform.
        var n = command.Numbers;
        var matrix = pageTransform.Multiply(new(n[0], n[1], n[2], n[3], n[4], n[5]));
        var size = Math.Abs(n[6]);
        var advance = command.Glyphs!.Sum(glyph => glyph.Advance) * n[7];
        var origin = matrix.Point(0, n[8] + size * .85);
        var right = matrix.Point(advance, n[8] + size * .85);
        var down = matrix.Point(0, n[8] - size * .35);
        var frame = Frame(origin, right, down);
        if (frame is null)
        {
            return;
        }
        if (state.TextMode is 1 or 2 or 5 or 6)
        {
            warn("Outlined text is approximated with filled editable text.");
        }
        var name = Regex.Replace(command.FontName ?? "Arial", "^[A-Z]{6}\\+", "");
        var bold = Regex.IsMatch(name, "bold|black|heavy", RegexOptions.IgnoreCase);
        var italic = Regex.IsMatch(name, "italic|oblique", RegexOptions.IgnoreCase);
        var family = Regex.Replace(name, "[-,]?(BoldItalic|BoldOblique|Bold|Italic|Oblique|Regular|Roman|MT)$", "", RegexOptions.IgnoreCase);
        family = Regex.Replace(family, "[-,]?(Bold|Italic|Oblique)$", "", RegexOptions.IgnoreCase);
        family = family.StartsWith("Helvetica", StringComparison.OrdinalIgnoreCase) ? "Arial" : family.StartsWith("Times", StringComparison.OrdinalIgnoreCase) ? "Times New Roman" : family.StartsWith("Courier", StringComparison.OrdinalIgnoreCase) ? "Courier New" : family;
        // Text has already been mapped to Unicode. Symbol-font codepoint mappings
        // must not be applied a second time by PowerPoint.
        if (Regex.IsMatch(family, "^(Wingdings|Webdings|Symbol)", RegexOptions.IgnoreCase)
            && text.EnumerateRunes().All(rune => rune.Value < 0xe000 || rune.Value > 0xf8ff))
        {
            family = "Arial";
        }
        var fontSize = Math.Clamp((int)Math.Round(size * Math.Sqrt(matrix.C * matrix.C + matrix.D * matrix.D) * 100), 100, 400000);
        var textGradient = state.TextMode is 1 or 5 ? state.StrokeGradient : state.FillGradient;
        var properties = new XElement(A + "rPr", new XAttribute("lang", "en-US"), new XAttribute("sz", fontSize), new XAttribute("b", bold ? 1 : 0), new XAttribute("i", italic ? 1 : 0), new XAttribute("dirty", 0),
            textGradient?.Fill(origin, right, down, alpha) ?? Fill(state.TextMode is 1 or 5 ? state.Stroke : state.Fill, alpha), new XElement(A + "latin", new XAttribute("typeface", family)), new XElement(A + "ea", new XAttribute("typeface", family)), new XElement(A + "cs", new XAttribute("typeface", family)));
        var start = matrix.Point(0, n[8]);
        var end = matrix.Point(advance, n[8]);
        var length = Math.Sqrt(Math.Pow(end.X - start.X, 2) + Math.Pow(end.Y - start.Y, 2));
        var height = size * Math.Sqrt(matrix.C * matrix.C + matrix.D * matrix.D);
        var fragment = new TextFragment(Clean(text), properties, frame, start, end,
            new((end.X - start.X) / length, (end.Y - start.Y) / length),
            new(matrix.C * size / height, matrix.D * size / height), height, textGradient, alpha);
        if (textLine.Count > 0 && !CanJoinText(fragment))
        {
            FlushText();
        }
        textLine.Add(fragment);
    }

    private void MatchOutlinedText(PdfCanvasCommand command)
    {
        // Some PDF producers paint gradient glyph outlines followed immediately by
        // transparent searchable text. Unmatched invisible/OCR text stays excluded.
        var outline = lastOutline;
        lastOutline = null;
        if (outline is null || command.Glyphs is not { Count: > 0 } glyphs)
        {
            return;
        }
        var n = command.Numbers;
        var matrix = pageTransform.Multiply(new(n[0], n[1], n[2], n[3], n[4], n[5]));
        var size = Math.Abs(n[6]);
        var visibleGlyphs = glyphs.ToList();
        while (visibleGlyphs.Count > 0 && string.IsNullOrWhiteSpace(visibleGlyphs[^1].Text))
        {
            visibleGlyphs.RemoveAt(visibleGlyphs.Count - 1);
        }
        if (visibleGlyphs.Count == 0)
        {
            return;
        }
        var width = visibleGlyphs.Sum(glyph => glyph.Advance) * n[7];
        var corners = new[] { matrix.Point(0, n[8] + size * .85), matrix.Point(width, n[8] + size * .85), matrix.Point(0, n[8] - size * .25), matrix.Point(width, n[8] - size * .25) };
        var left = corners.Min(point => point.X);
        var right = corners.Max(point => point.X);
        var top = corners.Min(point => point.Y);
        var bottom = corners.Max(point => point.Y);
        var tolerance = size * Math.Sqrt(matrix.C * matrix.C + matrix.D * matrix.D) * .4;
        if (Math.Abs(outline.Left - left) <= tolerance && Math.Abs(outline.Left + outline.Width - right) <= tolerance
            && Math.Abs(outline.Top - top) <= tolerance && Math.Abs(outline.Top + outline.Height - bottom) <= tolerance
            && outline.Width * outline.Height > (right - left) * (bottom - top) * .35)
        {
            outlinedText.Add(outline.Shape, (command, outline.PaintState with { TextMode = 0 }));
        }
    }

    private void RecoverOutlinedText()
    {
        var originals = new List<XElement>();
        void Finish()
        {
            FlushText();
            foreach (var original in originals)
            {
                original.Remove();
                ShapeCount--;
            }
            originals.Clear();
            textInsertionAnchor = null;
        }
        foreach (var shape in tree.Elements().ToArray())
        {
            if (!outlinedText.TryGetValue(shape, out var source))
            {
                Finish();
                continue;
            }
            textInsertionAnchor ??= shape;
            state = source.PaintState;
            Text(source.Command);
            originals.Add(shape);
        }
        Finish();
    }

    private static double Dot(PdfImagePoint point, PdfImagePoint axis) => point.X * axis.X + point.Y * axis.Y;

    private static PdfImagePoint Difference(PdfImagePoint a, PdfImagePoint b) => new(a.X - b.X, a.Y - b.Y);

    private bool CanJoinText(TextFragment next)
    {
        var first = textLine[0];
        var previous = textLine[^1];
        // Skewed, reflected, overlapping and reverse-order runs retain independent boxes.
        if (Dot(first.Direction, next.Direction) < .99999 || Dot(first.Up, next.Up) < .99999
            || Math.Abs(Dot(first.Direction, first.Up)) > .001
            || first.Direction.X * first.Up.Y - first.Direction.Y * first.Up.X >= 0)
        {
            return false;
        }
        var baselineGap = Math.Abs(Dot(Difference(next.Start, first.Start), first.Up));
        var minimumHeight = Math.Min(first.Height, next.Height);
        var maximumHeight = Math.Max(first.Height, next.Height);
        var sameBaseline = baselineGap <= minimumHeight * TextHeuristics.LineBaselineTolerance
            || minimumHeight / maximumHeight < TextHeuristics.SuperscriptSizeRatio
                && baselineGap <= maximumHeight * TextHeuristics.SuperscriptBaselineTolerance;
        var gap = Dot(Difference(next.Start, previous.End), first.Direction);
        return sameBaseline && gap >= -Math.Min(previous.Height, next.Height) * .08
            && gap <= Math.Max(previous.Height, next.Height) * TextHeuristics.TextColumnGapFactor;
    }

    private void FlushText()
    {
        if (textLine.Count == 0)
        {
            return;
        }
        var first = textLine[0];
        var frame = first.Frame;
        var reference = textLine.MaxBy(run => run.Height)!;
        if (textLine.Count > 1)
        {
            var top = textLine.Max(run => Dot(Difference(run.Start, first.Start), first.Up) + run.Height * .85);
            var bottom = textLine.Min(run => Dot(Difference(run.Start, first.Start), first.Up) - run.Height * .35);
            var width = Dot(Difference(textLine[^1].End, first.Start), first.Direction);
            PdfImagePoint Point(double x, double y) => new(first.Start.X + first.Direction.X * x + first.Up.X * y, first.Start.Y + first.Direction.Y * x + first.Up.Y * y);
            frame = Frame(Point(0, top), Point(width, top), Point(0, bottom))!;
        }
        var groupTop = textLine.Max(run => Dot(Difference(run.Start, first.Start), first.Up) + run.Height * .85);
        var groupBottom = textLine.Min(run => Dot(Difference(run.Start, first.Start), first.Up) - run.Height * .35);
        var groupWidth = Dot(Difference(textLine[^1].End, first.Start), first.Direction);
        PdfImagePoint GroupPoint(double x, double y) => new(first.Start.X + first.Direction.X * x + first.Up.X * y, first.Start.Y + first.Direction.Y * x + first.Up.Y * y);
        var paragraph = new XElement(A + "p", new XElement(A + "pPr", new XAttribute("marL", 0), new XAttribute("indent", 0), new XElement(A + "buNone")));
        for (var i = 0; i < textLine.Count; i++)
        {
            var run = textLine[i];
            var text = run.Text;
            if (run.Gradient is not null)
            {
                run.Properties.Element(A + "gradFill")!.ReplaceWith(run.Gradient.Fill(GroupPoint(0, groupTop), GroupPoint(groupWidth, groupTop), GroupPoint(0, groupBottom), run.Alpha));
            }
            if (i > 0 && Dot(Difference(run.Start, textLine[i - 1].End), first.Direction) > run.Height * TextHeuristics.WordGapFactor
                && !char.IsWhiteSpace(textLine[i - 1].Text[^1]) && !char.IsWhiteSpace(text[0]))
            {
                text = " " + text;
            }
            var offset = Dot(Difference(run.Start, reference.Start), first.Up);
            if (textLine.Count > 1 && Math.Abs(offset) > .01)
            {
                run.Properties.SetAttributeValue("baseline", (int)Math.Round(offset / run.Height * 100000));
            }
            paragraph.Add(new XElement(A + "r", run.Properties, new XElement(A + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text)));
        }
        var body = new XElement(P + "txBody", new XElement(A + "bodyPr", new XAttribute("wrap", "none"), new XAttribute("lIns", 0), new XAttribute("tIns", 0), new XAttribute("rIns", 0), new XAttribute("bIns", 0), new XAttribute("anchor", "t"), new XElement(A + "noAutofit")), new XElement(A + "lstStyle"), paragraph);
        var shape = new XElement(P + "sp", NonVisual("Text", true), new XElement(P + "spPr", frame, Rect(), new XElement(A + "noFill"), new XElement(A + "ln", new XElement(A + "noFill"))), body);
        if (textInsertionAnchor is not null)
        {
            textInsertionAnchor.AddBeforeSelf(shape);
        }
        else
        {
            tree.Add(shape);
        }
        TextCount++;
        textLine.Clear();
    }

    private void Picture(PdfCanvasCommand command, PdfMatrix matrix)
    {
        lastOutline = null;
        if (state.FillAlpha <= 0)
        {
            return;
        }
        var frame = Frame(matrix.Point(0, 1), matrix.Point(1, 1), matrix.Point(0, 0));
        if (frame is null)
        {
            return;
        }
        var bytes = System.Convert.FromBase64String(command.Data ?? throw new InvalidDataException("Missing image bytes."));
        var extension = command.Kind == "jpeg" ? "jpg" : "png";
        if (extension == "png")
        {
            bytes = PptxPng.Encode(bytes, command.Width, command.Height);
        }
        else if (command.AlphaData is not null)
        {
            bytes = PptxMaskedImage.Decode(command, bytes);
            extension = "png";
        }
        var id = "rId" + (relationships.Count + 2);
        relationships.Add(Relation(id, "image", package.AddMedia(bytes, extension)));
        tree.Add(new XElement(P + "pic", new XElement(P + "nvPicPr", new XElement(P + "cNvPr", new XAttribute("id", ++objectId), new XAttribute("name", "Image " + objectId)), new XElement(P + "cNvPicPr"), new XElement(P + "nvPr")),
            new XElement(P + "blipFill", new XElement(A + "blip", new XAttribute(R + "embed", id), new XElement(A + "alphaModFix", new XAttribute("amt", (int)Math.Round(state.FillAlpha * 100000)))), new XElement(A + "stretch", new XElement(A + "fillRect"))), new XElement(P + "spPr", frame, Rect())));
        ImageCount++;
    }

    private void Paint(string op, PdfMatrix matrix)
    {
        if (op != "n")
        {
            lastOutline = null;
        }
        var points = path.SelectMany(item => item.Points).ToArray();
        if (op == "n" || !options.IncludePaths || points.Length == 0)
        {
            path.Clear();
            return;
        }
        var left = points.Min(point => point.X);
        var top = points.Min(point => point.Y);
        var width = Math.Max(.01, points.Max(point => point.X) - left);
        var height = Math.Max(.01, points.Max(point => point.Y) - top);
        var geometry = new XElement(A + "path", new XAttribute("w", Emu(width)), new XAttribute("h", Emu(height)));
        foreach (var segment in path)
        {
            geometry.Add(new XElement(A + segment.Kind, segment.Points.Select(point => new XElement(A + "pt", new XAttribute("x", Emu(point.X - left)), new XAttribute("y", Emu(point.Y - top))))));
        }
        var fill = op is "f" or "F" or "f*" or "B" or "B*" or "b" or "b*";
        var stroke = op is "S" or "s" or "B" or "B*" or "b" or "b*";
        var frame = Frame(new(left, top), new(left + width, top), new(left, top + height));
        if (frame is not null)
        {
            if (op.EndsWith('*'))
            {
                warn("Even-odd path fills use PowerPoint's path fill behavior and may differ.");
            }
            var lineWidth = Math.Clamp(Emu(Math.Abs(state.LineWidth) * Math.Sqrt(Math.Abs(matrix.A * matrix.D - matrix.B * matrix.C))), 0, 20116800);
            tree.Add(new XElement(P + "sp", NonVisual("Shape", false), new XElement(P + "spPr", frame,
                new XElement(A + "custGeom", new XElement(A + "avLst"), new XElement(A + "gdLst"), new XElement(A + "ahLst"), new XElement(A + "cxnLst"), new XElement(A + "rect", new XAttribute("l", "l"), new XAttribute("t", "t"), new XAttribute("r", "r"), new XAttribute("b", "b")), new XElement(A + "pathLst", geometry)),
                fill ? state.FillGradient?.Fill(new(left, top), new(left + width, top), new(left, top + height), state.FillAlpha) ?? Fill(state.Fill, state.FillAlpha) : new XElement(A + "noFill"), new XElement(A + "ln", new XAttribute("w", lineWidth), stroke ? state.StrokeGradient?.Fill(new(left, top), new(left + width, top), new(left, top + height), state.StrokeAlpha) ?? Fill(state.Stroke, state.StrokeAlpha) : new XElement(A + "noFill")))));
            ShapeCount++;
            if (fill && !stroke && state.FillGradient is not null && state.FillAlpha > 0 && path.Any(segment => segment.Kind == "cubicBezTo"))
            {
                lastOutline = new(tree.Elements().Last(), state, left, top, width, height);
            }
        }
        path.Clear();
    }

    private XElement? Frame(PdfImagePoint origin, PdfImagePoint right, PdfImagePoint down)
    {
        var x = right.X - origin.X;
        var y = right.Y - origin.Y;
        var vx = down.X - origin.X;
        var vy = down.Y - origin.Y;
        var width = Math.Sqrt(x * x + y * y);
        var height = Math.Sqrt(vx * vx + vy * vy);
        if (!double.IsFinite(width + height) || width < .0001 || height < .0001)
        {
            return null;
        }
        var corners = new[] { origin, right, down, new PdfImagePoint(right.X + vx, right.Y + vy) };
        if (corners.Max(point => point.X) <= 0 || corners.Max(point => point.Y) <= 0 || corners.Min(point => point.X) >= slideWidth || corners.Min(point => point.Y) >= slideHeight)
        {
            return null;
        }
        if (Math.Abs((x * vx + y * vy) / (width * height)) > .001)
        {
            warn("Skewed text/pictures are approximated with rotated rectangular objects.");
        }
        var centerX = origin.X + (x + vx) / 2;
        var centerY = origin.Y + (y + vy) / 2;
        var angle = (Math.Atan2(y, x) * 180 / Math.PI + 360) % 360;
        return new(A + "xfrm", new XAttribute("rot", (int)Math.Round(angle * 60000) % 21600000), new XAttribute("flipV", x * vy - y * vx < 0 ? 1 : 0),
            new XElement(A + "off", new XAttribute("x", Emu(centerX - width / 2)), new XAttribute("y", Emu(centerY - height / 2))),
            new XElement(A + "ext", new XAttribute("cx", Emu(width)), new XAttribute("cy", Emu(height))));
    }

    private XElement NonVisual(string name, bool text) => new(P + "nvSpPr", new XElement(P + "cNvPr", new XAttribute("id", ++objectId), new XAttribute("name", name + " " + objectId)), new XElement(P + "cNvSpPr", new XAttribute("txBox", text ? 1 : 0)), new XElement(P + "nvPr"));

    private static XElement Rect() => new(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"));

    private static XElement Fill(string color, double alpha) => new(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", color), new XElement(A + "alpha", new XAttribute("val", (int)Math.Round(alpha * 100000)))));

    private static string Color(double[] values)
    {
        var rgb = values.Length == 1 ? new[] { values[0], values[0], values[0] } : values.Length == 4 ? values.Take(3).Select(value => 1 - Math.Min(1, value + values[3])).ToArray() : values;
        return string.Concat(rgb.Take(3).Select(value => ((int)Math.Round(Math.Clamp(value, 0, 1) * 255)).ToString("X2")));
    }

    private static string Clean(string text)
    {
        return string.Concat(text.EnumerateRunes().Where(rune => rune.Value > 0xffff || XmlConvert.IsXmlChar((char)rune.Value)).Select(rune => rune.ToString()));
    }
}
