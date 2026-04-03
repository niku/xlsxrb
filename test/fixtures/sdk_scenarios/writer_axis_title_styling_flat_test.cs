// Validates that a writer-generated XLSX has axis titles with font and spPr (fill + line).
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    if (!xml.Contains("<c:catAx>") || !xml.Contains("<c:valAx>"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing catAx or valAx");
    }
    else if (!xml.Contains("b=\"1\""))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing bold on axis title font");
    }
    else if (!xml.Contains("val=\"EEEEFF\""))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing fill color EEEEFF on cat axis title");
    }
    else if (!xml.Contains("<a:noFill/>"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing noFill on val axis title");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
