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
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text("hello"))
            }
        )
        { RowIndex = 1, Height = 25.0, CustomHeight = true },
        new Row() { RowIndex = 3, Hidden = true }
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
