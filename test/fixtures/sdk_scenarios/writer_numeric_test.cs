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

    Cell FindCell(string reference) =>
        worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference?.Value == reference)
            ?? throw new Exception($"{reference} cell is missing.");

    var a1 = FindCell("A1");
    if (a1.DataType?.Value != null)
    {
        throw new Exception($"A1 should have no DataType (numeric) but got '{a1.DataType?.Value}'.");
    }
    var a1Value = a1.CellValue?.Text ?? throw new Exception("A1 has no value.");
    if (a1Value != "42")
    {
        throw new Exception($"Expected A1 to be '42' but got '{a1Value}'.");
    }

    var b1 = FindCell("B1");
    if (b1.DataType?.Value != null)
    {
        throw new Exception($"B1 should have no DataType (numeric) but got '{b1.DataType?.Value}'.");
    }
    var b1Value = b1.CellValue?.Text ?? throw new Exception("B1 has no value.");
    if (b1Value != "3.14")
    {
        throw new Exception($"Expected B1 to be '3.14' but got '{b1Value}'.");
    }
}
finally
{
    document.Dispose();
}
