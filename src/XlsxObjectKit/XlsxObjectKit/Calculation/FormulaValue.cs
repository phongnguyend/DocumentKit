using System.Globalization;

namespace XlsxObjectKit.Calculation;

internal readonly record struct FormulaValue(string Kind, string Text, double Number = 0)
{
    public static FormulaValue Numeric(double value) => double.IsFinite(value) ? new("n", value.ToString("R", CultureInfo.InvariantCulture), value) : Error("#NUM!");

    public static FormulaValue Error(string value) => new("e", value);

    public static FormulaValue Pending(string reason) => new("pending", reason);

    public static FormulaValue Boolean(bool value) => new("b", value ? "TRUE" : "FALSE", value ? 1 : 0);

    public bool Failed => Kind is "e" or "pending";

    public FormulaValue AsNumber()
    {
        if (Failed)
        {
            return this;
        }
        if (Kind is "n" or "b" or "blank")
        {
            return Numeric(Number);
        }
        return double.TryParse(Text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? Numeric(number) : Error("#VALUE!");
    }

    public FormulaValue AsBoolean()
    {
        if (Kind == "str" && bool.TryParse(Text, out var boolean))
        {
            return Boolean(boolean);
        }
        var number = AsNumber();
        return number.Failed ? number : Boolean(number.Number != 0);
    }
}
