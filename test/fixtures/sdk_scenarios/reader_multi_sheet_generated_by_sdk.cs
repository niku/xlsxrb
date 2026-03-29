var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    // Sheet1
    var worksheetPart1 = workbookPart.AddNewPart<WorksheetPart>();
    var sheetData1 = new SheetData(
        new Row(
            new Cell
            {
                CellReference = "A1",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text("main"))
            }
        )
    );
    worksheetPart1.Worksheet = new Worksheet(sheetData1);

    // Data sheet
    var worksheetPart2 = workbookPart.AddNewPart<WorksheetPart>();
    var sheetData2 = new SheetData(
        new Row(
            new Cell
            {
                CellReference = "A1",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text("data"))
            }
        )
    );
    worksheetPart2.Worksheet = new Worksheet(sheetData2);

    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
    sheets.Append(new Sheet
    {
        Id = workbookPart.GetIdOfPart(worksheetPart1),
        SheetId = 1,
        Name = "Sheet1"
    });
    sheets.Append(new Sheet
    {
        Id = workbookPart.GetIdOfPart(worksheetPart2),
        SheetId = 2,
        Name = "Data"
    });

    workbookPart.Workbook.Save();
    worksheetPart1.Worksheet.Save();
    worksheetPart2.Worksheet.Save();
}
finally
{
    document.Dispose();
}
