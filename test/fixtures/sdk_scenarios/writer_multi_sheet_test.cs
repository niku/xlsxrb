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
    var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList()
        ?? throw new Exception("Sheets element is missing.");

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

    if (sheets.Count != 2)
    {
        throw new Exception($"Expected 2 sheets but got {sheets.Count}.");
    }

    if (sheets[0].Name?.Value != "Sheet1")
    {
        throw new Exception($"Expected first sheet name 'Sheet1' but got '{sheets[0].Name?.Value}'.");
    }
    if (sheets[1].Name?.Value != "Data")
    {
        throw new Exception($"Expected second sheet name 'Data' but got '{sheets[1].Name?.Value}'.");
    }

    // Check Sheet1 A1
    var ws1 = (WorksheetPart)workbookPart.GetPartById(sheets[0].Id!.Value!);
    var a1Sheet1 = ws1.Worksheet.Descendants<Cell>()
        .FirstOrDefault(c => c.CellReference?.Value == "A1")
        ?? throw new Exception("Sheet1 A1 cell is missing.");
    var val1 = ReadSharedString(a1Sheet1, "Sheet1 A1");
    if (val1 != "main")
    {
        throw new Exception($"Expected Sheet1 A1='main' but got '{val1}'.");
    }

    // Check Data A1
    var ws2 = (WorksheetPart)workbookPart.GetPartById(sheets[1].Id!.Value!);
    var a1Sheet2 = ws2.Worksheet.Descendants<Cell>()
        .FirstOrDefault(c => c.CellReference?.Value == "A1")
        ?? throw new Exception("Data A1 cell is missing.");
    var val2 = ReadSharedString(a1Sheet2, "Data A1");
    if (val2 != "data")
    {
        throw new Exception($"Expected Data A1='data' but got '{val2}'.");
    }
}
finally
{
    document.Dispose();
}
