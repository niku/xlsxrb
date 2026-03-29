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

    if (firstSheet.Id?.Value is null)
    {
        throw new Exception("Sheet relationship ID is missing.");
    }

    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    Cell FindCell(string reference) =>
        worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => c.CellReference?.Value == reference)
            ?? throw new Exception($"{reference} cell is missing.");

    string ReadInlineString(Cell cell, string reference)
    {
        if (cell.DataType?.Value != CellValues.InlineString)
        {
            throw new Exception($"{reference} must be stored as inline string.");
        }

        return cell.InlineString?.Text?.Text
            ?? cell.InlineString?.InnerText
            ?? throw new Exception($"{reference} inline string value is missing.");
    }

    var a1Value = ReadInlineString(FindCell("A1"), "A1");
    if (a1Value != "hello")
    {
        throw new Exception($"Expected A1 to be 'hello' but got '{a1Value}'.");
    }

    var b1Value = ReadInlineString(FindCell("B1"), "B1");
    if (b1Value != "world")
    {
        throw new Exception($"Expected B1 to be 'world' but got '{b1Value}'.");
    }
}
finally
{
    document.Dispose();
}
