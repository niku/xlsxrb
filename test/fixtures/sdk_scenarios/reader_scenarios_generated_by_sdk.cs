var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook(
        new Sheets(new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" })
    );

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");

    var scenarios = new Scenarios { Current = 0U, Show = 0U };
    var scenario = new Scenario
    {
        Name = "Best Case",
        User = "Admin",
        Comment = "Optimistic",
        Count = 2U
    };
    scenario.Append(new InputCells { CellReference = "A1", Val = "200" });
    scenario.Append(new InputCells { CellReference = "B1", Val = "300" });
    scenarios.Append(scenario);

    worksheetPart.Worksheet = new Worksheet(
        new SheetData(
            new Row(new Cell { CellReference = "A1", DataType = CellValues.Number, CellValue = new CellValue("100") })
            { RowIndex = 1 }
        ),
        scenarios
    );

    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
    stylesPart.Stylesheet = new Stylesheet(
        new Fonts(new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" })),
        new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
        new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())),
        new CellFormats(new CellFormat { FontId = 0, FillId = 0, BorderId = 0 })
    );

    document.Save();
}
finally
{
    document.Dispose();
}
