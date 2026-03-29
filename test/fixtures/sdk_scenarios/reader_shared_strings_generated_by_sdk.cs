var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var sharedStringsPart = workbookPart.AddNewPart<SharedStringTablePart>();
    sharedStringsPart.SharedStringTable = new SharedStringTable(
        new SharedStringItem(new Text("hello")),
        new SharedStringItem(new Text("world"))
    );

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    var sheetData = new SheetData(
        new Row(
            new Cell
            {
                CellReference = "A1",
                DataType = CellValues.SharedString,
                CellValue = new CellValue("0")
            },
            new Cell
            {
                CellReference = "B1",
                DataType = CellValues.SharedString,
                CellValue = new CellValue("1")
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

    sharedStringsPart.SharedStringTable.Save();
    workbookPart.Workbook.Save();
    worksheetPart.Worksheet.Save();
}
finally
{
    document.Dispose();
}
