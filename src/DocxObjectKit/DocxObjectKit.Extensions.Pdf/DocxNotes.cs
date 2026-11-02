using System.Xml.Linq;
using static DocxObjectKit.Models.DocxNamespaces;

namespace DocxObjectKit.Extensions.Pdf;

public sealed partial class DocxLayout
{
    private sealed record Note(string Key, int Number, string Part, XElement Content);
    private readonly Dictionary<string, Note> notes = new();
    private Dictionary<int, double> noteReservations = new();

    private double FlowLimit => result.Height - bottom - (measuringTable ? 0 : noteReservations.GetValueOrDefault(page.Number));

    private void LoadNotes()
    {
        foreach (var kind in new[] { "footnote", "endnote" })
        {
            var part = "word/" + kind + "s.xml";
            if (!document.PartNames.Contains(part))
            {
                continue;
            }
            var definitions = document.ReadXml(part).Root?.Elements(W + kind).Where(item => Value(item, "type") is null or "normal")
                .GroupBy(item => Value(item, "id") ?? "").ToDictionary(group => group.Key, group => group.First());
            var number = 0;
            foreach (var reference in xml.Descendants(W + (kind + "Reference")))
            {
                var id = Value(reference, "id") ?? "";
                var key = kind + ":" + id;
                if (notes.ContainsKey(key))
                {
                    continue;
                }
                if (definitions?.TryGetValue(id, out var definition) == true)
                {
                    var content = new XElement(definition);
                    number++;
                    foreach (var marker in content.Descendants(W + (kind + "Ref")).ToArray())
                    {
                        marker.ReplaceWith(new XElement(W + "t", "[" + number + "] "));
                    }
                    var ordinal = 0;
                    foreach (var paragraph in content.Descendants(W + "p"))
                    {
                        paragraphIds[paragraph] = "note-" + key + "-" + ordinal++;
                    }
                    ordinal = 0;
                    foreach (var picture in content.Descendants().Where(item => item.Name == WP + "inline" || item.Name == WP + "anchor"))
                    {
                        imageIds[picture] = "note-" + key + "-image-" + ordinal++;
                    }
                    notes.Add(key, new(key, number, part, content));
                }
            }
        }
        if (notes.Count > 0 && xml.Descendants().Any(item => item.Name == W + "numRestart" || item.Name == W + "numFmt"))
        {
            warnings.Add("Notes currently use continuous decimal numbering; custom note numbering is not implemented.");
        }
    }

    public DocxLayoutResult Build(CancellationToken cancellationToken = default) => BuildWithNotes(0, cancellationToken);

    private DocxLayoutResult BuildWithNotes(int pass, CancellationToken cancellationToken)
    {
        BuildBody(cancellationToken);
        var footnotes = new Dictionary<int, (DocxPage Content, double Height)>();
        var overflow = new List<Note>();
        foreach (var target in result.Pages.ToArray())
        {
            var keys = target.Runs.Select(run => run.NoteKey).Where(key => key?.StartsWith("footnote:", StringComparison.Ordinal) == true).Distinct().ToArray();
            if (keys.Length == 0)
            {
                continue;
            }
            if (sections.Length > 0)
            {
                ApplySection(sections[pageSections.GetValueOrDefault(target)]);
            }
            var scratch = new DocxPage(0, "notes", "", [], [], [], []);
            var savedPage = page;
            var savedY = y;
            page = scratch;
            result.Pages.Add(scratch);
            y = 8;
            measuringTable = true;
            try
            {
                foreach (var key in keys)
                {
                    RenderNote(notes[key!]);
                }
                if (y > (result.Height - top - bottom) / 2)
                {
                    overflow.AddRange(keys.Select(key => notes[key!]));
                    warnings.Add("Long footnotes are placed after the document; split footnote continuation is not yet supported.");
                }
                else
                {
                    footnotes[target.Number] = (scratch, y);
                }
            }
            finally
            {
                measuringTable = false;
                result.Pages.Remove(scratch);
                page = savedPage;
                y = savedY;
                currentPart = "word/document.xml";
            }
        }
        var reservation = new Dictionary<int, double>(noteReservations);
        foreach (var (number, measured) in footnotes)
        {
            reservation[number] = Math.Max(reservation.GetValueOrDefault(number), measured.Height);
        }
        if (pass < 6 && reservation.Any(item => item.Value > noteReservations.GetValueOrDefault(item.Key) + .01))
        {
            using var retry = new DocxLayout(document) { noteReservations = reservation };
            return retry.BuildWithNotes(pass + 1, cancellationToken);
        }
        if (pass == 6 && reservation.Any(item => item.Value > noteReservations.GetValueOrDefault(item.Key) + .01))
        {
            throw new NotSupportedException("Footnote pagination did not converge within seven layout passes.");
        }
        foreach (var target in result.Pages)
        {
            if (!footnotes.TryGetValue(target.Number, out var measured))
            {
                continue;
            }
            if (sections.Length > 0)
            {
                ApplySection(sections[pageSections.GetValueOrDefault(target)]);
            }
            var shift = result.Height - bottom - measured.Height;
            target.Rules.Add(new(left, shift, Math.Min(144, result.Width - left - right), 0, StrokeWidth: .5));
            target.Runs.AddRange(measured.Content.Runs.Select(run => run with { Y = run.Y + shift, Baseline = run.Baseline + shift }));
            target.Images.AddRange(measured.Content.Images.Select(image => image with { Y = image.Y + shift, ReadOnly = true }));
            target.Rules.AddRange(measured.Content.Rules.Select(rule => rule with { Y = rule.Y + shift }));
        }
        var endnotes = notes.Values.Where(note => note.Key.StartsWith("endnote:", StringComparison.Ordinal)).Concat(overflow).ToArray();
        if (endnotes.Length > 0)
        {
            if (sections.Length > 0)
            {
                ApplySection(sections[^1]);
            }
            page = NewPage();
            foreach (var note in endnotes)
            {
                cancellationToken.ThrowIfCancellationRequested();
                RenderNote(note);
            }
            currentPart = "word/document.xml";
        }
        if (noteReservations.Keys.Any(number => !footnotes.ContainsKey(number)))
        {
            warnings.Add("Footnote pagination retains conservative spacing where references moved to another page.");
        }
        RenderPageDecorations();
        var labels = new Dictionary<int, int>();
        var formattedLabels = new Dictionary<int, string>();
        var displayed = 0;
        var lastSection = -1;
        foreach (var target in result.Pages)
        {
            var section = pageSections.GetValueOrDefault(target);
            displayed++;
            if (section != lastSection && sections.Length > section && int.TryParse(Value(sections[section].Element(W + "pgNumType"), "start"), out var start) && start >= 0)
            {
                displayed = start;
            }
            labels[target.Number] = displayed;
            formattedLabels[target.Number] = FormatNumber(displayed, sections.Length > section ? Value(sections[section].Element(W + "pgNumType"), "fmt") ?? "decimal" : "decimal");
            lastSection = section;
        }
        for (var index = 0; index < result.Pages.Count; index++)
        {
            var target = result.Pages[index];
            target.TextObjects.RemoveAll(item => item.Id.StartsWith("note-", StringComparison.Ordinal));
            result.Pages[index] = target with { Title = target.TextObjects.FirstOrDefault()?.Text ?? "Page " + (index + 1) };
        }
        return result with { Warnings = warnings.Order(StringComparer.Ordinal).ToArray(), PageNumbers = labels, PageLabels = formattedLabels };
    }

    private void RenderNote(Note note)
    {
        currentPart = note.Part;
        previousParagraphStyle = null;
        pendingAfter = 0;
        foreach (var block in Blocks(note.Content))
        {
            if (block.Name == W + "p")
            {
                Paragraph(block, left, result.Width - left - right, true);
            }
            else if (block.Name == W + "tbl")
            {
                Table(block, left, result.Width - left - right);
            }
        }
    }
}
