var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
    stylesPart.Stylesheet = new Stylesheet(
        new Fonts(new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" })) { Count = 1 },
        new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
        ) { Count = 2 },
        new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())) { Count = 1 },
        new CellStyleFormats(new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }) { Count = 1 },
        new CellFormats(
            new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 }
        ) { Count = 1 },
        new CellStyles(new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }) { Count = 1 }
    );
    stylesPart.Stylesheet.Save();

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    var sheetData = new SheetData(
        new Row(
            new Cell { CellReference = "A1", DataType = CellValues.InlineString,
                       InlineString = new InlineString(new Text("data")) }
        )
    );
    var pageSetup = new PageSetup
    {
        PageOrder = PageOrderValues.OverThenDown,
        BlackAndWhite = true,
        Draft = true,
        CellComments = CellCommentsValues.AtEnd,
        FirstPageNumber = 5,
        UseFirstPageNumber = true,
        HorizontalDpi = 300,
        VerticalDpi = 300,
        Orientation = OrientationValues.Landscape,
        PaperSize = 9
    };

    worksheetPart.Worksheet = new Worksheet(sheetData, pageSetup);

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
