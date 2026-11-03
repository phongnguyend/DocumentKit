using System.Text.Json;
using XlsxObjectKit.Application;
using XlsxObjectKit.Editors;
using XlsxObjectKit.Writers;

try
{
    if (args is ["new", var destination])
    {
        File.WriteAllBytes(destination, XlsxWriter.Write(XlsxEditor.Create()));
    }
    else if (args is ["inspect", var input])
    {
        Console.WriteLine(JsonSerializer.Serialize(WorkbookApplication.Inspect(File.ReadAllBytes(input)), new JsonSerializerOptions { WriteIndented = true }));
    }
    else if (args is ["calculate", var calculationInput, var calculationOutput])
    {
        var result = WorkbookApplication.EditWithDiagnostics(File.ReadAllBytes(calculationInput), new("recalculate", ""));
        File.WriteAllBytes(calculationOutput, result.Bytes);
        foreach (var warning in result.Warnings)
        {
            Console.Error.WriteLine(warning);
        }
    }
    else if (args is ["edit", var source, var edits, var output])
    {
        var request = JsonSerializer.Deserialize<XlsxEditRequest>(File.ReadAllText(edits), new JsonSerializerOptions(JsonSerializerDefaults.Web)) ?? throw new ArgumentException("Missing edit JSON.");
        File.WriteAllBytes(output, WorkbookApplication.Edit(File.ReadAllBytes(source), request));
    }
    else
    {
        Console.Error.WriteLine("Usage: new output.xlsx | inspect input.xlsx | edit input.xlsx edit.json output.xlsx | calculate input.xlsx output.xlsx");
        return 2;
    }
    return 0;
}
catch (Exception exception) when (exception is IOException or ArgumentException or NotSupportedException or JsonException or InvalidOperationException)
{
    Console.Error.WriteLine(exception.Message);
    return 1;
}
