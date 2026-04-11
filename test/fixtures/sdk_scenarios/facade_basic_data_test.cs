// Validates that an XLSX with basic cell data passes OpenXML validation.
// This validates the fundamental XLSX structure produced by the Facade API.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList()
        ?? throw new Exception("Sheets collection is missing.");
    if (sheets.Count == 0)
        throw new Exception("No sheets found.");

    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheets[0].Id.Value);
    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault()
        ?? throw new Exception("SheetData is missing.");

    var rows = sheetData.Elements<Row>().ToList();
    if (rows.Count < 1)
        throw new Exception("Expected at least 1 row of data.");

    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).ToList();
    if (errors.Count > 0)
    {
        var messages = string.Join("\n", errors.Select(e => $"  - {e.Description} (Part: {e.Part?.Uri}, Path: {e.Path?.XPath})"));
        throw new Exception($"OpenXML Validation errors:\n{messages}");
    }
}
finally
{
    document.Dispose();
}
