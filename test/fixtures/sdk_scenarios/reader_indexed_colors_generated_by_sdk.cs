var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook(
        new Sheets(new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" })
    );

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
    worksheetPart.Worksheet = new Worksheet(
        new SheetData(
            new Row(new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("test") })
            { RowIndex = 1 }
        )
    );

    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
    var indexedColors = new IndexedColors();
    indexedColors.Append(new RgbColor { Rgb = "FF000000" });
    indexedColors.Append(new RgbColor { Rgb = "FFFFFFFF" });
    indexedColors.Append(new RgbColor { Rgb = "FFFF0000" });
    var colors = new Colors();
    colors.Append(indexedColors);

    stylesPart.Stylesheet = new Stylesheet(
        new Fonts(new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" })),
        new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
        new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())),
        new CellFormats(new CellFormat { FontId = 0, FillId = 0, BorderId = 0 }),
        colors
    );

    document.Save();
}
finally
{
    document.Dispose();
}
