var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var ws1 = workbookPart.AddNewPart<WorksheetPart>();
    ws1.Worksheet = new Worksheet(new SheetData(
        new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("hello")) })
    ));

    var ws2 = workbookPart.AddNewPart<WorksheetPart>();
    ws2.Worksheet = new Worksheet(new SheetData());

    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(ws1), SheetId = 1, Name = "Sheet1" });
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(ws2), SheetId = 2, Name = "Data" });

    var definedNames = workbookPart.Workbook.AppendChild(new DefinedNames());
    definedNames.Append(new DefinedName("Sheet1!$A$1:$B$10") { Name = "MyRange" });
    definedNames.Append(new DefinedName("Data!$C$1") { Name = "LocalName", LocalSheetId = 1U });
    definedNames.Append(new DefinedName("42") { Name = "HiddenConst", Hidden = true });

    workbookPart.Workbook.Save();
    ws1.Worksheet.Save();
    ws2.Worksheet.Save();
}
finally
{
    document.Dispose();
}
