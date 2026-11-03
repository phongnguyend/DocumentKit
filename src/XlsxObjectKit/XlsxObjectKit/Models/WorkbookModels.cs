using System.Text.RegularExpressions;

namespace XlsxObjectKit.Models;

public sealed record XlsxImage(string Id, string Name, string DataUrl, double X, double Y, double Width, double Height, bool Editable);

public sealed record XlsxCell(string Address, string Value, string Type, string? Formula, int Style)
{
    public bool HasCachedValue { get; init; }
}

public sealed record XlsxBorder(string Style, string Color);

public sealed record XlsxStyle(bool Bold, bool Italic, string Color, string Background, string Alignment, string NumberFormat)
{
    public int NumberFormatId { get; init; }

    public string FontName { get; init; } = "Calibri";

    public double FontSize { get; init; } = 11;

    public string Underline { get; init; } = "none";

    public bool Strike { get; init; }

    public bool WrapText { get; init; }

    public string VerticalAlignment { get; init; } = "bottom";

    public int Indent { get; init; }

    public int TextRotation { get; init; }

    public int ReadingOrder { get; init; }

    public IReadOnlyDictionary<string, XlsxBorder> Borders { get; init; } = new Dictionary<string, XlsxBorder>();
}

public sealed record XlsxSheet(string Id, string Name, string Part, bool Hidden, IReadOnlyList<XlsxCell> Cells,
    IReadOnlyList<string> Merges, IReadOnlyDictionary<int, double> RowHeights, IReadOnlyDictionary<int, double> ColumnWidths)
{
    public string? TabColor { get; init; }

    public IReadOnlyList<XlsxObjectKit.Editors.XlsxConditionalRule> ConditionalRules { get; init; } = [];

    public IReadOnlyList<int> HiddenRows { get; init; } = [];

    public IReadOnlyList<int> HiddenColumns { get; init; } = [];

    public IReadOnlyDictionary<int, int> RowStyles { get; init; } = new Dictionary<int, int>();

    public IReadOnlyDictionary<int, int> ColumnStyles { get; init; } = new Dictionary<int, int>();

    public IReadOnlyList<XlsxImage> Images { get; init; } = [];

    public double DefaultRowHeight { get; init; } = 15;

    public double DefaultColumnWidth { get; init; } = 8.43;
}

public sealed record XlsxWorkbook(IReadOnlyList<XlsxSheet> Sheets, IReadOnlyList<XlsxStyle> Styles, IReadOnlyList<string> Warnings)
{
    public bool Date1904 { get; init; }
}

public static class CellAddress
{
    public static (int Row, int Column) Parse(string address)
    {
        if (!Regex.IsMatch(address, "^[A-Za-z]{1,3}[1-9][0-9]{0,6}$"))
        {
            throw new ArgumentException("Use an A1 cell address.");
        }
        var column = 0;
        var index = 0;
        while (char.IsLetter(address[index]))
        {
            column = column * 26 + char.ToUpperInvariant(address[index++]) - 'A' + 1;
        }
        var row = int.Parse(address[index..]);
        if (row > 1048576 || column > 16384)
        {
            throw new ArgumentException("Cell is outside Excel worksheet limits.");
        }
        return (row, column);
    }

    public static string Format(int row, int column)
    {
        var letters = "";
        while (column > 0)
        {
            column--;
            letters = (char)('A' + column % 26) + letters;
            column /= 26;
        }
        return letters + row;
    }
}
