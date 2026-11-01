using System.Xml.Linq;
using PdfObjectKit.Interpreters;
using PdfObjectKit.Models;
using static PdfObjectKit.Converters.PptxPackage;

namespace PdfObjectKit.Converters;

// An axial PDF gradient remains anchored in page space, even across text runs.
internal sealed class PptxGradient
{
    private readonly double[] values;
    private readonly double x;
    private readonly double y;
    private readonly double offset;

    internal PptxGradient(double[] values, PdfMatrix pageTransform)
    {
        if (values.Length < 18 || (values.Length - 10) % 4 != 0)
        {
            throw new InvalidDataException("Invalid axial gradient stops.");
        }
        this.values = values;
        var matrix = pageTransform.Multiply(new(values[4], values[5], values[6], values[7], values[8], values[9]));
        var dx = values[2] - values[0];
        var dy = values[3] - values[1];
        var lengthSquared = dx * dx + dy * dy;
        var determinant = matrix.A * matrix.D - matrix.B * matrix.C;
        if (lengthSquared <= 1e-12 || Math.Abs(determinant) <= 1e-12)
        {
            throw new NotSupportedException("Degenerate axial gradient.");
        }
        x = (dx * matrix.D - dy * matrix.B) / (lengthSquared * determinant);
        y = (-dx * matrix.C + dy * matrix.A) / (lengthSquared * determinant);
        offset = -x * matrix.E - y * matrix.F - (dx * values[0] + dy * values[1]) / lengthSquared;
    }

    internal XElement Fill(PdfImagePoint origin, PdfImagePoint right, PdfImagePoint down, double alpha)
    {
        var start = x * origin.X + y * origin.Y + offset;
        var dx = x * (right.X - origin.X) + y * (right.Y - origin.Y);
        var dy = x * (down.X - origin.X) + y * (down.Y - origin.Y);
        var minimum = start + Math.Min(0, dx) + Math.Min(0, dy);
        var maximum = start + Math.Max(0, dx) + Math.Max(0, dy);
        var width = Math.Sqrt(Math.Pow(right.X - origin.X, 2) + Math.Pow(right.Y - origin.Y, 2));
        var height = Math.Sqrt(Math.Pow(down.X - origin.X, 2) + Math.Pow(down.Y - origin.Y, 2));
        var angle = (Math.Atan2(dy / Math.Max(height, .001), dx / Math.Max(width, .001)) * 180 / Math.PI + 360) % 360;
        var positions = new SortedSet<double> { 0, 1 };
        if (maximum - minimum > 1e-12)
        {
            for (var i = 10; i < values.Length; i += 4)
            {
                var position = (values[i] - minimum) / (maximum - minimum);
                if (position > 0 && position < 1)
                {
                    positions.Add(position);
                }
            }
        }
        var stops = new XElement(A + "gsLst");
        foreach (var position in positions)
        {
            var t = minimum + position * (maximum - minimum);
            var index = 10;
            while (index + 4 < values.Length && values[index + 4] < t)
            {
                index += 4;
            }
            var next = Math.Min(index + 4, values.Length - 4);
            var fraction = next == index ? 0 : Math.Clamp((t - values[index]) / (values[next] - values[index]), 0, 1);
            var color = string.Concat(Enumerable.Range(1, 3).Select(channel => ((int)Math.Round(Math.Clamp(values[index + channel] + fraction * (values[next + channel] - values[index + channel]), 0, 1) * 255)).ToString("X2")));
            stops.Add(new XElement(A + "gs", new XAttribute("pos", (int)Math.Round(position * 100000)), new XElement(A + "srgbClr", new XAttribute("val", color), new XElement(A + "alpha", new XAttribute("val", (int)Math.Round(alpha * 100000))))));
        }
        return new XElement(A + "gradFill", new XAttribute("rotWithShape", 1), stops, new XElement(A + "lin", new XAttribute("ang", (int)Math.Round(angle * 60000) % 21600000), new XAttribute("scaled", 0)));
    }
}
