using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: dotnet run --project vendor/sdk_runner -- <scenario.cs> <xlsx-path>");
    return 2;
}

var scenarioPath = Path.GetFullPath(args[0]);
var xlsxPath = Path.GetFullPath(args[1]);

if (!File.Exists(scenarioPath))
{
    Console.Error.WriteLine($"Scenario file not found: {scenarioPath}");
    return 2;
}

if (!File.Exists(xlsxPath))
{
    Console.Error.WriteLine($"XLSX file not found: {xlsxPath}");
    return 2;
}

var code = await File.ReadAllTextAsync(scenarioPath);
var globals = new ScenarioContext(xlsxPath);

var options = ScriptOptions.Default
    .AddReferences(
        typeof(object).Assembly,
        typeof(Enumerable).Assembly,
        typeof(Console).Assembly,
        typeof(SpreadsheetDocument).Assembly,
        typeof(OpenXmlValidator).Assembly
    )
    .AddImports(
        "System",
        "System.IO",
        "System.Linq",
        "System.Collections.Generic",
        "DocumentFormat.OpenXml",
        "DocumentFormat.OpenXml.Packaging",
        "DocumentFormat.OpenXml.Spreadsheet",
        "DocumentFormat.OpenXml.Validation"
    );

try
{
    await CSharpScript.RunAsync(code, options, globals, typeof(ScenarioContext));
    Console.WriteLine("SCENARIO_PASS");
    return 0;
}
catch (CompilationErrorException ex)
{
    Console.Error.WriteLine("SCENARIO_COMPILE_ERROR");
    foreach (var diagnostic in ex.Diagnostics)
    {
        Console.Error.WriteLine(diagnostic.ToString());
    }
    return 1;
}
catch (Exception ex)
{
    Console.Error.WriteLine("SCENARIO_FAIL");
    Console.Error.WriteLine(ex.ToString());
    return 1;
}

public sealed class ScenarioContext
{
    public ScenarioContext(string xlsxPath)
    {
        XlsxPath = xlsxPath;
    }

    public string XlsxPath { get; }
}
