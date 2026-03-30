var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook(new Sheets(
        new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }
    ));

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
    var sheetView = new SheetView
    {
        ShowZeros = false,
        View = SheetViewValues.PageBreakPreview,
        ShowOutlineSymbols = false,
        ShowRuler = false,
        WorkbookViewId = 0
    };
    var sheetViews = new SheetViews();
    sheetViews.Append(sheetView);

    worksheetPart.Worksheet = new Worksheet(
        sheetViews,
        new SheetData(
            new Row(new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("test") })
            { RowIndex = 1 }
        )
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
