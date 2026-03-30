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
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
            new Fill(new PatternFill { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Rgb = "FFFF0000" } })
        ) { Count = 3 },
        new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())) { Count = 1 },
        new CellStyleFormats(new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }) { Count = 1 },
        new CellFormats(new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 }) { Count = 1 },
        new CellStyles(new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }) { Count = 1 },
        new DifferentialFormats(
            new DifferentialFormat(
                new Fill(new PatternFill { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor { Rgb = "FFFF0000" } })
            )
        ) { Count = 1 }
    );
    stylesPart.Stylesheet.Save();

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    var autoFilter = new AutoFilter { Reference = "A1:B5" };
    autoFilter.Append(new FilterColumn { ColumnId = 0 });
    autoFilter.Elements<FilterColumn>().First().Append(
        new ColorFilter { FormatId = 0U }
    );
    autoFilter.Append(new FilterColumn { ColumnId = 1 });
    autoFilter.Elements<FilterColumn>().Last().Append(
        new IconFilter { IconSet = IconSetValues.ThreeArrows, IconId = 1U }
    );

    var sheetData = new SheetData(
        new Row(
            new Cell { CellReference = "A1", CellValue = new CellValue("H1"), DataType = CellValues.InlineString, InlineString = new InlineString(new Text("H1")) },
            new Cell { CellReference = "B1", CellValue = new CellValue("H2"), DataType = CellValues.InlineString, InlineString = new InlineString(new Text("H2")) }
        )
    );

    worksheetPart.Worksheet = new Worksheet(sheetData, autoFilter);

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
