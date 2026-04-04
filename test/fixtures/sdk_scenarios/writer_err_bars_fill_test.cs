// Validates that a writer-generated XLSX with error bar fill properties is valid.
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    if (!xml.Contains("<c:errBars>"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:errBars> element");
    }
    else if (!xml.Contains("<a:solidFill>"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing solidFill in errBars spPr");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
