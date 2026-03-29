var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var validationErrors = validator.Validate(document).Take(10).ToList();
    if (validationErrors.Any())
    {
        var message = string.Join(Environment.NewLine, validationErrors.Select(e => e.Description));
        throw new Exception($"OpenXmlValidator reported errors:{Environment.NewLine}{message}");
    }

    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet definition is missing.");

    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id!.Value!);
    var rows = worksheetPart.Worksheet.Descendants<Row>().ToList();

    // Row 1 should have custom height 25.0
    var row1 = rows.FirstOrDefault(r => r.RowIndex?.Value == 1)
        ?? throw new Exception("Row 1 is missing.");
    if (row1.CustomHeight?.Value != true)
    {
        throw new Exception("Row 1 should have customHeight=true.");
    }
    if (Math.Abs(row1.Height!.Value - 25.0) > 0.01)
    {
        throw new Exception($"Expected row 1 height=25.0 but got {row1.Height.Value}.");
    }

    // Row 3 should be hidden
    var row3 = rows.FirstOrDefault(r => r.RowIndex?.Value == 3)
        ?? throw new Exception("Row 3 is missing.");
    if (row3.Hidden?.Value != true)
    {
        throw new Exception("Row 3 should be hidden.");
    }
}
finally
{
    document.Dispose();
}
