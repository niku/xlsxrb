var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook(
        new Sheets(new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" })
    );

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");

    var dataRefs = new DataReferences { Count = 2 };
    dataRefs.Append(new DataReference { Reference = "A1:B10", Sheet = "Sheet1" });
    dataRefs.Append(new DataReference { Reference = "C1:D10", Name = "Range2" });

    var dc = new DataConsolidate
    {
        Function = DataConsolidateFunctionValues.Average,
        StartLabels = true,
        Link = true
    };
    dc.Append(dataRefs);

    worksheetPart.Worksheet = new Worksheet(
        new SheetData(
            new Row(new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("data") })
            { RowIndex = 1 }
        ),
        dc
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
