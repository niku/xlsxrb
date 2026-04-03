// Validates that a writer-generated XLSX has chart title with font + spPr (fill + line).
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    if (!xml.Contains("<c:title>"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:title> element");
    }
    else if (!xml.Contains("b=\"1\""))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing bold attribute on rPr");
    }
    else if (!xml.Contains("<a:solidFill>") || !xml.Contains("val=\"FFFF00\""))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing fill color FFFF00");
    }
    else if (!xml.Contains("<a:prstDash val=\"dash\""))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing line dash");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
