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
    var worksheetPart = workbookPart.WorksheetParts.First();
    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>() ?? throw new Exception("SheetData is missing.");
    var row = sheetData.Elements<Row>().First();
    var cells = row.Elements<Cell>().ToList();

    if (cells.Count < 7)
        throw new Exception($"Expected at least 7 cells but got {cells.Count}.");

    Action<Cell, string, string> assertErrorCell = (cell, expectedRef, expectedError) =>
    {
        if (cell.CellReference?.Value != expectedRef)
            throw new Exception($"Expected cell ref {expectedRef} but got {cell.CellReference?.Value}.");
        if (cell.DataType?.Value != CellValues.Error)
            throw new Exception($"Cell {expectedRef}: expected DataType=Error but got {cell.DataType?.Value}.");
        if (cell.CellValue?.Text != expectedError)
            throw new Exception($"Cell {expectedRef}: expected value '{expectedError}' but got '{cell.CellValue?.Text}'.");
    };

    assertErrorCell(cells[0], "A1", "#N/A");
    assertErrorCell(cells[1], "B1", "#DIV/0!");
    assertErrorCell(cells[2], "C1", "#VALUE!");
    assertErrorCell(cells[3], "D1", "#REF!");
    assertErrorCell(cells[4], "E1", "#NAME?");
    assertErrorCell(cells[5], "F1", "#NUM!");
    assertErrorCell(cells[6], "G1", "#NULL!");
}
finally
{
    document.Dispose();
}
