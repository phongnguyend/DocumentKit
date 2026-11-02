using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using SkiaSharp;
using static PptxObjectKit.Graphs.PptxPackageGraph;

namespace PptxObjectKit.Extensions.Pdf;

internal sealed partial class SlideRenderer
{
    private readonly Dictionary<string, SKTypeface> fonts = new(StringComparer.OrdinalIgnoreCase);
    private SKTypeface? fallbackFont;

    private SKTypeface Font(XElement properties)
    {
        var name = (string?)C(properties, "latin")?.Attribute("typeface") ?? "+mn-lt";
        if (name.StartsWith("+", StringComparison.Ordinal))
        {
            var family = name.StartsWith("+mj", StringComparison.Ordinal) ? "majorFont" : "minorFont";
            name = (string?)C(D(scene.Theme, family).FirstOrDefault(), "latin")?.Attribute("typeface") ?? "Arial";
        }
        var bold = N(properties, "b") != 0;
        var italic = N(properties, "i") != 0;
        var key = name + "/" + bold + "/" + italic;
        if (fonts.TryGetValue(key, out var cached))
        {
            return cached;
        }
        var style = new SKFontStyle(bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal, SKFontStyleWidth.Normal, italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright);
        var face = SKFontManager.Default.MatchFamily(name, style);
        if (face is null || !string.Equals(face.FamilyName, name, StringComparison.OrdinalIgnoreCase))
        {
            face?.Dispose();
            if (fallbackFont is null)
            {
                using var stream = typeof(SlideRenderer).Assembly.GetManifestResourceStream("PptxObjectKit.Extensions.Pdf.Fonts.NotoSans-Regular.ttf")!;
                using var data = SKData.Create(stream);
                fallbackFont = SKTypeface.FromData(data) ?? throw new InvalidOperationException("Cannot load the bundled fallback font.");
            }
            Warn($"Font '{name}' is unavailable; Noto Sans was substituted. Install the original font for closer text layout.");
            return fallbackFont;
        }
        fonts.Add(key, face);
        return face;
    }

    private sealed record TextPiece(string Text, XElement Style, float Size, float Width);

    private sealed class TextLine
    {
        internal List<TextPiece> Pieces { get; } = [];

        internal float Width { get; set; }

        internal float Height { get; set; }

        internal float Ascent { get; set; }

        internal float X { get; set; }

        internal float Y { get; set; }

        internal string Alignment { get; set; } = "l";
    }

    private void Text(SKCanvas canvas, XElement body, float width, float height)
    {
        var bodyPr = C(body, "bodyPr");
        var vertical = (string?)bodyPr?.Attribute("vert") ?? "horz";
        if (vertical != "horz")
        {
            Warn("Vertical text is not yet supported.");
            return;
        }
        var left = N(bodyPr, "lIns", 91440) / Emu;
        var right = N(bodyPr, "rIns", 91440) / Emu;
        var top = N(bodyPr, "tIns", 45720) / Emu;
        var bottom = N(bodyPr, "bIns", 45720) / Emu;
        var available = Math.Max(1, width - left - right);
        var wrap = (string?)bodyPr?.Attribute("wrap") != "none";
        var fontScale = N(C(bodyPr, "normAutofit"), "fontScale", 100000) / 100000;
        if (fontScale <= 0 || fontScale > 1)
        {
            throw new InvalidDataException("Text autofit scale must be greater than zero and no more than 100 percent.");
        }
        var lineSpacingReduction = N(C(bodyPr, "normAutofit"), "lnSpcReduction");
        if (lineSpacingReduction < 0 || lineSpacingReduction > 100000)
        {
            throw new InvalidDataException("Autofit line-spacing reduction must be between zero and 100 percent.");
        }
        var lines = new List<TextLine>();
        float y = 0;
        var counters = new Dictionary<int, int>();
        foreach (var paragraph in body.Elements(A + "p"))
        {
            cancellation.ThrowIfCancellationRequested();
            var pPr = C(paragraph, "pPr");
            var baseStyle = C(pPr, "defRPr") ?? new XElement(A + "rPr", new XAttribute("sz", "1800"));
            var baseSize = Math.Clamp(N(baseStyle, "sz", 1800) / 100 * fontScale, .1f, 1000);
            var margin = N(pPr, "marL") / Emu;
            var indent = N(pPr, "indent") / Emu;
            y += Spacing(C(pPr, "spcBef"), baseSize, 0);
            var line = new TextLine { X = margin + indent, Y = y, Height = baseSize * 1.2f, Alignment = (string?)pPr?.Attribute("algn") ?? "l" };
            var pendingWord = new List<(string Text, XElement Style)>();
            var committingWord = false;
            var bullet = C(pPr, "buChar");
            var auto = C(pPr, "buAutoNum");
            if ((bullet is not null || auto is not null) && D(paragraph, "t").Any(item => !string.IsNullOrWhiteSpace(item.Value)))
            {
                var bulletText = ((string?)bullet?.Attribute("char") ?? "\u2022") + " ";
                if (auto is not null)
                {
                    var level = (int)N(pPr, "lvl");
                    foreach (var deeper in counters.Keys.Where(key => key > level).ToArray())
                    {
                        counters.Remove(deeper);
                    }
                    var number = (int)N(auto, "startAt", counters.GetValueOrDefault(level) + 1);
                    counters[level] = number;
                    var type = (string?)auto.Attribute("type") ?? "arabicPeriod";
                    var value = type.StartsWith("alpha", StringComparison.Ordinal) ? ((char)((type.StartsWith("alphaUc", StringComparison.Ordinal) ? 'A' : 'a') + (number - 1) % 26)).ToString() : number.ToString(CultureInfo.InvariantCulture);
                    bulletText = type.EndsWith("ParenBoth", StringComparison.Ordinal) ? "(" + value + ") " : type.EndsWith("ParenR", StringComparison.Ordinal) ? value + ") " : type.EndsWith("Plain", StringComparison.Ordinal) ? value + " " : value + ". ";
                    if (type.StartsWith("roman", StringComparison.Ordinal))
                    {
                        Warn("Roman list numbering uses Arabic numerals.");
                    }
                }
                if (((string?)C(pPr, "buFont")?.Attribute("typeface"))?.Contains("Wingdings", StringComparison.OrdinalIgnoreCase) == true && bulletText.Trim() is "\u00a7" or "\uf0a7")
                {
                    bulletText = "\u25aa ";
                }
                if (bulletText.Any(character => character >= 0xe000 && character <= 0xf8ff))
                {
                    bulletText = "\u2022 ";
                }
                var bulletStyle = new XElement(C(C(paragraph, "r"), "rPr") ?? baseStyle);
                if (C(pPr, "buClr") is { } bulletColor)
                {
                    bulletStyle.Elements(A + "solidFill").Remove();
                    bulletStyle.Add(new XElement(A + "solidFill", bulletColor.Elements().Select(item => new XElement(item))));
                }
                bulletStyle.SetAttributeValue("u", "none");
                bulletStyle.SetAttributeValue("strike", "noStrike");
                Add(bulletText, bulletStyle);
                if (line.Pieces.Count > 0 && -indent > line.Width)
                {
                    var padding = -indent - line.Width;
                    line.Pieces[^1] = line.Pieces[^1] with { Width = line.Pieces[^1].Width + padding };
                    line.Width += padding;
                }
            }
            foreach (var run in paragraph.Elements())
            {
                if (run.Name.LocalName == "br")
                {
                    CommitWord();
                    Finish();
                    continue;
                }
                if (run.Name.LocalName is not ("r" or "fld"))
                {
                    continue;
                }
                var style = C(run, "rPr") ?? baseStyle;
                var text = C(run, "t")?.Value ?? "";
                if (string.Equals((string?)C(style, "sym")?.Attribute("typeface"), "Wingdings", StringComparison.OrdinalIgnoreCase))
                {
                    // Symbol runs can also contain ordinary Latin text. Map only known PUA symbols.
                    text = text.Replace('\uf0e8', '\u2794').Replace('\uf0a7', '\u25aa')
                        .Replace('\uf0fc', '\u2713').Replace('\uf0fb', '\u2717').Replace('\uf04c', '\u2639');
                }
                if (Regex.IsMatch(text, @"[\u0590-\u08ff\u0900-\u0dff]"))
                {
                    Warn("Complex script shaping is not supported; text may require a shaping-capable fallback converter.");
                }
                if ((string?)style.Attribute("cap") is "all" or "small")
                {
                    text = text.ToUpperInvariant();
                }
                foreach (var token in Regex.Split(text.Replace("\t", "    "), @"([ \r\n]+)").Where(item => item.Length > 0))
                {
                    if (string.IsNullOrWhiteSpace(token) && !token.Contains('\u00a0'))
                    {
                        CommitWord();
                        if (token.Contains('\n'))
                        {
                            Finish();
                        }
                        else
                        {
                            Add(token, style);
                        }
                    }
                    else
                    {
                        pendingWord.Add((token, style));
                    }
                }
            }
            CommitWord();
            Finish();
            y += Spacing(C(pPr, "spcAft"), baseSize, 0);

            void CommitWord()
            {
                if (pendingWord.Count == 0)
                {
                    return;
                }
                var total = 0f;
                foreach (var fragment in pendingWord)
                {
                    var size = Math.Clamp(N(fragment.Style, "sz", baseSize / fontScale * 100) / 100 * fontScale, .1f, 1000);
                    using var font = MakeFont(fragment.Style, size);
                    total += font.MeasureText(fragment.Text);
                }
                committingWord = total <= Math.Max(1, available - margin);
                if (wrap && committingWord && line.Pieces.Count > 0 && line.X + line.Width + total > available)
                {
                    Finish();
                }
                foreach (var fragment in pendingWord)
                {
                    Add(fragment.Text, fragment.Style);
                }
                committingWord = false;
                pendingWord.Clear();
            }

            void Add(string text, XElement style, bool resolveGlyphs = true)
            {
                var size = Math.Clamp(N(style, "sz", baseSize / fontScale * 100) / 100 * fontScale, .1f, 1000);
                using var font = MakeFont(style, size);
                if (resolveGlyphs && !font.ContainsGlyphs(text))
                {
                    var normal = new StringBuilder();
                    foreach (var rune in text.EnumerateRunes())
                    {
                        if (font.ContainsGlyph(rune.Value))
                        {
                            normal.Append(rune.ToString());
                            continue;
                        }
                        if (normal.Length > 0)
                        {
                            Add(normal.ToString(), style, false);
                            normal.Clear();
                        }
                        using var fallback = GlyphFallback(rune.Value);
                        var glyphStyle = new XElement(style);
                        if (fallback is not null)
                        {
                            glyphStyle.Elements(A + "latin").Remove();
                            glyphStyle.Add(new XElement(A + "latin", new XAttribute("typeface", fallback.FamilyName)));
                        }
                        else
                        {
                            Warn($"No available font contains character U+{rune.Value:X4}; its glyph cannot be rendered.");
                        }
                        Add(rune.ToString(), glyphStyle, false);
                    }
                    if (normal.Length > 0)
                    {
                        Add(normal.ToString(), style, false);
                    }
                    return;
                }
                var measured = font.MeasureText(text);
                if (wrap && measured > Math.Max(1, available - margin) && new StringInfo(text).LengthInTextElements > 1)
                {
                    var elements = StringInfo.GetTextElementEnumerator(text);
                    while (elements.MoveNext())
                    {
                        Add(elements.GetTextElement(), style);
                    }
                    return;
                }
                if (wrap && !committingWord && line.Pieces.Count > 0 && line.X + line.Width + measured > available && !string.IsNullOrWhiteSpace(text))
                {
                    Finish();
                }
                if (line.Pieces.Count == 0 && string.IsNullOrWhiteSpace(text))
                {
                    return;
                }
                line.Pieces.Add(new(text, style, size, measured));
                line.Width += measured;
                line.Height = Math.Max(line.Height, font.Metrics.Descent - font.Metrics.Ascent);
                line.Ascent = Math.Max(line.Ascent, -font.Metrics.Ascent);
            }

            void Finish()
            {
                if (line.Ascent == 0)
                {
                    line.Ascent = baseSize;
                }
                lines.Add(line);
                var spacing = C(pPr, "lnSpc");
                // Autofit subtracts percentage points from proportional line spacing.
                // Fixed point spacing and paragraph before/after spacing are independent.
                y += C(spacing, "spcPts") is not null
                    ? Spacing(spacing, line.Height, line.Height)
                    : line.Height * Math.Max(0, N(C(spacing, "spcPct"), "val", 100000) - lineSpacingReduction) / 100000;
                line = new TextLine { X = margin, Y = y, Height = baseSize * 1.2f, Alignment = (string?)pPr?.Attribute("algn") ?? "l" };
            }
        }
        var innerHeight = Math.Max(0, height - top - bottom);
        var anchor = (string?)bodyPr?.Attribute("anchor") ?? "t";
        var offset = anchor == "ctr" ? (innerHeight - y) / 2 : anchor == "b" ? innerHeight - y : 0;
        if (N(bodyPr, "numCol", 1) > 1)
        {
            Warn("Multicolumn text boxes are rendered as a single column.");
        }
        canvas.Save();
        try
        {
            if ((string?)bodyPr?.Attribute("vertOverflow") == "clip")
            {
                canvas.ClipRect(new(0, 0, width, height));
            }
            foreach (var line in lines)
            {
                var x = left + line.X;
                if (line.Alignment == "ctr")
                {
                    x += (available - line.X - line.Width) / 2;
                }
                else if (line.Alignment == "r")
                {
                    x += available - line.X - line.Width;
                }
                else if (line.Alignment is "just" or "dist")
                {
                    Warn("Justified text uses left alignment.");
                }
                foreach (var piece in line.Pieces)
                {
                    using var font = MakeFont(piece.Style, piece.Size);
                    using var paint = Fill(piece.Style, width, height) ?? new SKPaint
                    {
                        IsAntialias = true,
                        Color = C(piece.Style, "noFill") is not null ? SKColors.Transparent : SKColors.Black
                    };
                    var baseline = top + offset + line.Y + line.Ascent - N(piece.Style, "baseline") / 100000 * piece.Size;
                    canvas.DrawText(piece.Text, x, baseline, SKTextAlign.Left, font, paint);
                    if ((string?)piece.Style.Attribute("u") is { } underline && underline != "none")
                    {
                        paint.StrokeWidth = Math.Max(.5f, piece.Size / 16);
                        canvas.DrawLine(x, baseline + piece.Size * .1f, x + piece.Width, baseline + piece.Size * .1f, paint);
                    }
                    if ((string?)piece.Style.Attribute("strike") is { } strike && strike != "noStrike")
                    {
                        paint.StrokeWidth = Math.Max(.5f, piece.Size / 16);
                        canvas.DrawLine(x, baseline - piece.Size * .3f, x + piece.Width, baseline - piece.Size * .3f, paint);
                    }
                    x += piece.Width;
                }
            }
        }
        finally
        {
            canvas.Restore();
        }
    }

    private static SKTypeface? GlyphFallback(int character)
    {
        // Prefer text symbols over the platform emoji font for Wingdings equivalents.
        if (character is 0x2639 or 0x2794 or 0x25aa or 0x2713 or 0x2717)
        {
            var symbols = SKFontManager.Default.MatchFamily("Segoe UI Symbol");
            if (symbols is not null)
            {
                using var font = new SKFont(symbols);
                if (font.ContainsGlyph(character))
                {
                    return symbols;
                }
                symbols.Dispose();
            }
        }
        return SKFontManager.Default.MatchCharacter(character);
    }

    private SKFont MakeFont(XElement style, float size)
    {
        var face = Font(style);
        return new SKFont(face, size)
        {
            Embolden = N(style, "b") != 0 && !face.IsBold,
            SkewX = N(style, "i") != 0 && !face.IsItalic ? -.2f : 0
        };
    }

    private static float Spacing(XElement? spacing, float size, float fallback)
    {
        if (C(spacing, "spcPts") is { } points)
        {
            return N(points, "val") / 100;
        }
        return C(spacing, "spcPct") is { } percent ? N(percent, "val") / 100000 * size : fallback;
    }

    public void Dispose()
    {
        foreach (var font in fonts.Values)
        {
            font.Dispose();
        }
        fallbackFont?.Dispose();
    }
}
