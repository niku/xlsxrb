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
            },
            new Cell
            {
                CellReference = "B1",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text("Page"))
            },
            new Cell
            {
                CellReference = "C1",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text("Internal"))
            }
        )
    );

    worksheetPart.Worksheet = new Worksheet(sheetData);

    // HL1: A1 - external with display and tooltip
    var hlRel1 = worksheetPart.AddHyperlinkRelationship(
        new System.Uri("https://example.com"), true);

    // HL2: B1 - external with location
    var hlRel2 = worksheetPart.AddHyperlinkRelationship(
        new System.Uri("https://example.com/page"), true);

    var hyperlinks = new Hyperlinks(
        new Hyperlink
        {
            Reference = "A1",
            Id = hlRel1.Id,
            Display = "Example Site",
            Tooltip = "Click to visit"
        },
        new Hyperlink
        {
            Reference = "B1",
            Id = hlRel2.Id,
            Location = "Sheet2!A1"
        },
        new Hyperlink
        {
            Reference = "C1",
            Location = "Sheet1!D1"
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
