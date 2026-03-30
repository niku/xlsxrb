var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook(
        new Sheets(new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" })
    );

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");

    var ignoredErrors = new IgnoredErrors();
    var error = new IgnoredError
    {
        NumberStoredAsText = true,
        FormulaRange = true
    };
    error.SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("A1") });
    ignoredErrors.Append(error);

    worksheetPart.Worksheet = new Worksheet(
        new SheetData(
            new Row(new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("123") })
            { RowIndex = 1 }
        ),
        ignoredErrors
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
