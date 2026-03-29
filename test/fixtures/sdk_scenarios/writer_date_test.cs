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
    var a1Cell = worksheetPart.Worksheet.Descendants<Cell>()
        .FirstOrDefault(c => c.CellReference?.Value == "A1")
        ?? throw new Exception("A1 cell is missing.");

    // Jan 15, 2024 = serial 45306
    var expectedSerial = 45306;
    var actualValue = int.Parse(a1Cell.CellValue!.Text);
    if (actualValue != expectedSerial)
    {
        throw new Exception($"Expected serial {expectedSerial} but got {actualValue}.");
    }

    // Verify cell has a style
    if (a1Cell.StyleIndex == null)
    {
        throw new Exception("A1 should have a style index for date formatting.");
    }
}
finally
{
    document.Dispose();
}
