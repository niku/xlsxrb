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
    var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable
        ?? throw new Exception("SharedStringTable is missing.");
    var a1Cell = worksheetPart.Worksheet.Descendants<Cell>()
        .FirstOrDefault(c => c.CellReference?.Value == "A1")
        ?? throw new Exception("A1 cell is missing.");

    if (a1Cell.DataType?.Value != CellValues.SharedString)
    {
        throw new Exception("A1 must be stored as shared string.");
    }

    var sharedIndex = int.Parse(a1Cell.CellValue?.Text
        ?? throw new Exception("A1 shared string index is missing."));
    var actualValue = sharedStringTable.Elements<SharedStringItem>().ElementAt(sharedIndex).InnerText;

    if (actualValue != "hello")
    {
        throw new Exception($"Expected A1 to be 'hello' but got '{actualValue}'.");
    }
}
finally
{
    document.Dispose();
}
