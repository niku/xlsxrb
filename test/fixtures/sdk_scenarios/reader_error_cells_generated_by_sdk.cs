var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    var sheetData = new SheetData(
        new Row(
            new Cell
            {
                CellReference = "A1",
                DataType = CellValues.Error,
                CellValue = new CellValue("#DIV/0!")
            },
            new Cell
            {
                CellReference = "B1",
                DataType = CellValues.Error,
                CellValue = new CellValue("#VALUE!")
            }
        )
    );

    worksheetPart.Worksheet = new Worksheet(sheetData);

    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
    sheets.Append(new Sheet
    {
        Id = workbookPart.GetIdOfPart(worksheetPart),
        SheetId = 1,
        Name = "Sheet1"
    });

    workbookPart.Workbook.Save();
    worksheetPart.Worksheet.Save();
}
finally
{
    document.Dispose();
}
