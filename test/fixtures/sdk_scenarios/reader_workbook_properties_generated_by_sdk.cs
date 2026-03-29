var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    workbookPart.Workbook.WorkbookProperties = new WorkbookProperties { Date1904 = false, DefaultThemeVersion = 166925U };
    workbookPart.Workbook.BookViews = new BookViews(new WorkbookView { ActiveTab = 1U, FirstSheet = 0U });
    workbookPart.Workbook.CalculationProperties = new CalculationProperties { CalculationId = 191029U, FullCalculationOnLoad = true };

    var worksheetPart1 = workbookPart.AddNewPart<WorksheetPart>();
    worksheetPart1.Worksheet = new Worksheet(new SheetData(
        new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("main")) })
    ));

    var worksheetPart2 = workbookPart.AddNewPart<WorksheetPart>();
    worksheetPart2.Worksheet = new Worksheet(new SheetData(
        new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("data")) })
    ));

    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart1), SheetId = 1, Name = "Sheet1" });
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart2), SheetId = 2, Name = "Data" });

    workbookPart.Workbook.Save();
    worksheetPart1.Worksheet.Save();
    worksheetPart2.Worksheet.Save();
}
finally
{
    document.Dispose();
}
