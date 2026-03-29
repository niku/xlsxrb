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
    var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable
        ?? throw new Exception("SharedStringTable is missing.");
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

    string ReadSharedString(Cell cell, string reference)
    {
        if (cell.DataType?.Value != CellValues.SharedString)
        {
            throw new Exception($"{reference} must be stored as shared string.");
        }

        var sharedIndex = int.Parse(cell.CellValue?.Text
            ?? throw new Exception($"{reference} shared string index is missing."));
        return sharedStringTable.Elements<SharedStringItem>().ElementAt(sharedIndex).InnerText;
    }

    var a1Value = ReadSharedString(FindCell("A1"), "A1");
    if (a1Value != "hello")
    {
        throw new Exception($"Expected A1 to be 'hello' but got '{a1Value}'.");
    }

    var b1Value = ReadSharedString(FindCell("B1"), "B1");
    if (b1Value != "world")
    {
        throw new Exception($"Expected B1 to be 'world' but got '{b1Value}'.");
    }
}
finally
{
    document.Dispose();
}
