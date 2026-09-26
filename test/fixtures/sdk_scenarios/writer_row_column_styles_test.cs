var doc = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var validationErrors = validator.Validate(doc).Take(10).ToList();
    if (validationErrors.Any())
    {
        var message = string.Join(Environment.NewLine, validationErrors.Select(e => e.Description));
        throw new Exception($"OpenXmlValidator reported errors:{Environment.NewLine}{message}");
    }

    var workbookPart = doc.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var stylesPart = workbookPart.WorkbookStylesPart ?? throw new Exception("WorkbookStylesPart is missing.");
    var stylesheet = stylesPart.Stylesheet ?? throw new Exception("Stylesheet is missing.");

    // Verify cellXfs has style with alignment and protection
    var cellXfs = stylesheet.CellFormats ?? throw new Exception("CellFormats missing.");
    if (cellXfs.Count?.Value < 2)
        throw new Exception($"Expected at least 2 cellXfs, got {cellXfs.Count?.Value ?? 0}");

    var sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!.Value!);
    var worksheet = worksheetPart.Worksheet ?? throw new Exception("Worksheet is missing.");

    // Verify row style
    var row = worksheet.Descendants<Row>().FirstOrDefault()
        ?? throw new Exception("Row is missing.");
    if (row.StyleIndex?.Value != 1)
        throw new Exception($"Expected Row StyleIndex 1, got {row.StyleIndex?.Value}");
    if (row.CustomFormat?.Value != true)
        throw new Exception("Expected Row CustomFormat true");

    // Verify column style
    var col = worksheet.Descendants<Column>().FirstOrDefault()
        ?? throw new Exception("Column is missing.");
    if (col.Style?.Value != 1)
        throw new Exception($"Expected Column Style 1, got {col.Style?.Value}");
}
finally
{
    doc.Dispose();
}
