var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var ws1 = workbookPart.AddNewPart<WorksheetPart>();
    ws1.Worksheet = new Worksheet(new SheetData(
        new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("main")) })
    ));

    var ws2 = workbookPart.AddNewPart<WorksheetPart>();
    ws2.Worksheet = new Worksheet(new SheetData());

    var ws3 = workbookPart.AddNewPart<WorksheetPart>();
    ws3.Worksheet = new Worksheet(new SheetData());

    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(ws1), SheetId = 1, Name = "Sheet1" });
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(ws2), SheetId = 2, Name = "Hidden", State = SheetStateValues.Hidden });
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(ws3), SheetId = 3, Name = "VeryHidden", State = SheetStateValues.VeryHidden });

    workbookPart.Workbook.Save();
    ws1.Worksheet.Save();
    ws2.Worksheet.Save();
    ws3.Worksheet.Save();
}
finally
{
    document.Dispose();
}
