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

    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    var a3 = worksheetPart.Worksheet.Descendants<Cell>()
        .FirstOrDefault(c => c.CellReference?.Value == "A3")
        ?? throw new Exception("A3 cell is missing.");

    var formula = a3.CellFormula?.Text ?? throw new Exception("A3 has no formula.");
    if (formula != "SUM(A1:A2)")
    {
        throw new Exception($"Expected A3 formula 'SUM(A1:A2)' but got '{formula}'.");
    }

    var cachedValue = a3.CellValue?.Text ?? throw new Exception("A3 has no cached value.");
    if (cachedValue != "30")
    {
        throw new Exception($"Expected A3 cached value '30' but got '{cachedValue}'.");
    }
}
finally
{
    document.Dispose();
}
