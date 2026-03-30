var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook(
        new Sheets(new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" })
    );

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");

    var protectedRanges = new ProtectedRanges();
    var range1 = new ProtectedRange { Name = "EditArea" };
    range1.SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("A1:B10") });
    protectedRanges.Append(range1);

    var range2 = new ProtectedRange
    {
        Name = "SecureRange",
        AlgorithmName = "SHA-512",
        HashValue = "abc123",
        SaltValue = "salt456",
        SpinCount = 100000
    };
    range2.SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("C1:D5") });
    protectedRanges.Append(range2);

    worksheetPart.Worksheet = new Worksheet(
        new SheetData(
            new Row(new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("test") })
            { RowIndex = 1 }
        ),
        protectedRanges
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
