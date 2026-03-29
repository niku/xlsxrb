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
                InlineString = new InlineString(new Text("Example"))
            }
        )
    );

    worksheetPart.Worksheet = new Worksheet(sheetData);

    // Add hyperlink relationship
    var hyperlinkRelationship = worksheetPart.AddHyperlinkRelationship(
        new System.Uri("https://example.com"), true);

    var hyperlinks = new Hyperlinks(
        new Hyperlink
        {
            Reference = "A1",
            Id = hyperlinkRelationship.Id
        }
    );

    worksheetPart.Worksheet.AppendChild(hyperlinks);

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
