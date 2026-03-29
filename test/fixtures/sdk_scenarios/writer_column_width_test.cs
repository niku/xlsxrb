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
    var columns = worksheetPart.Worksheet.GetFirstChild<Columns>();
    if (columns == null)
    {
        throw new Exception("Columns element is missing.");
    }

    var colElements = columns.Elements<Column>().ToList();
    if (colElements.Count != 2)
    {
        throw new Exception($"Expected 2 col elements but got {colElements.Count}.");
    }

    // Column A (min=1, max=1, width=20)
    var colA = colElements[0];
    if (colA.Min?.Value != 1 || colA.Max?.Value != 1)
    {
        throw new Exception($"Expected col min=1 max=1 but got min={colA.Min?.Value} max={colA.Max?.Value}.");
    }
    if (Math.Abs(colA.Width!.Value - 20.0) > 0.01)
    {
        throw new Exception($"Expected column A width=20.0 but got {colA.Width.Value}.");
    }

    // Column C (min=3, max=3, width=15.5)
    var colC = colElements[1];
    if (colC.Min?.Value != 3 || colC.Max?.Value != 3)
    {
        throw new Exception($"Expected col min=3 max=3 but got min={colC.Min?.Value} max={colC.Max?.Value}.");
    }
    if (Math.Abs(colC.Width!.Value - 15.5) > 0.01)
    {
        throw new Exception($"Expected column C width=15.5 but got {colC.Width.Value}.");
    }
}
finally
{
    document.Dispose();
}
