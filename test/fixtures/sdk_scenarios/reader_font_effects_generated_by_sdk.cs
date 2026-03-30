var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook(new Sheets(
        new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }
    ));

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
    worksheetPart.Worksheet = new Worksheet(new SheetData(
        new Row(new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("effects"), StyleIndex = 1U })
        { RowIndex = 1 }
    ));

    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
    stylesPart.Stylesheet = new Stylesheet(
        new Fonts(
            new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" }),
            new Font(
                new Bold(),
                new Shadow(),
                new Outline(),
                new Condense(),
                new Extend(),
                new FontSize { Val = 14 },
                new FontName { Val = "Arial" }
            )
        ),
        new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
        ),
        new Borders(
            new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())
        ),
        new CellFormats(
            new CellFormat { FontId = 0, FillId = 0, BorderId = 0 },
            new CellFormat { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true }
        )
    );

    document.Save();
}
finally
{
    document.Dispose();
}
