var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    var sheetData = new SheetData(
        new Row(
            new Cell { CellReference = "A1", CellValue = new CellValue("10") }
        ),
        new Row(
            new Cell { CellReference = "A2", CellValue = new CellValue("20") }
        ),
        new Row(
            new Cell
            {
                CellReference = "A3",
                CellFormula = new CellFormula("SUM(A1:A2)"),
                CellValue = new CellValue("30")
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
