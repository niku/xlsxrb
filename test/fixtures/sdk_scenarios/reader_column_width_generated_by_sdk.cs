var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

    var columns = new Columns(
        new Column { Min = 1, Max = 1, Width = 20.0, CustomWidth = true },
        new Column { Min = 3, Max = 3, Width = 15.5, CustomWidth = true }
    );

    var sheetData = new SheetData(
        new Row(
            new Cell
            {
                CellReference = "A1",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text("hello"))
            }
        )
    );

    worksheetPart.Worksheet = new Worksheet(columns, sheetData);

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
