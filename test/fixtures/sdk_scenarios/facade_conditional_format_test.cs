using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).ToList();
    if (errors.Count > 0)
    {
        var messages = string.Join("\n", errors.Select(e => $"  - {e.Description} (Part: {e.Part?.Uri})"));
        throw new Exception($"OpenXML Validation errors:\n{messages}");
    }

    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart missing");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);

    var cfList = wsPart.Worksheet.Elements<ConditionalFormatting>().ToList();
    if (cfList.Count < 1)
        throw new Exception("SCENARIO_FAIL: ConditionalFormatting elements missing");

    var rule = cfList[0].Elements<ConditionalFormattingRule>().FirstOrDefault();
    if (rule == null)
        throw new Exception("SCENARIO_FAIL: ConditionalFormattingRule missing");

    if (rule.Type?.Value != ConditionalFormatValues.CellIs)
        throw new Exception($"SCENARIO_FAIL: Rule type expected cellIs, got {rule.Type?.Value}");

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
