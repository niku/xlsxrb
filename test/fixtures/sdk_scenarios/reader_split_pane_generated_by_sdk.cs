var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    var sheetViews = new SheetViews(
        new SheetView(
            new Pane
            {
                HorizontalSplit = 3000D,
                VerticalSplit = 2000D,
                TopLeftCell = "D5",
                ActivePane = PaneValues.BottomRight
            }
        ) { WorkbookViewId = 0 }
    );

    var sheetData = new SheetData(
        new Row(new Cell { CellReference = "A1", CellValue = new CellValue("split"), DataType = CellValues.InlineString, InlineString = new InlineString(new Text("split")) })
    );

    worksheetPart.Worksheet = new Worksheet(sheetViews, sheetData);

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
