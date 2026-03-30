var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook(
        new Sheets(new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" })
    );

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");

    var autoFilter = new AutoFilter { Reference = "A1:B10" };

    var filters = new Filters { CalendarType = CalendarValues.Gregorian };
    filters.Append(new Filter { Val = "Alice" });
    filters.Append(new DateGroupItem
    {
        DateTimeGrouping = DateTimeGroupingValues.Year,
        Year = 2024,
        Month = 1,
        Day = 1,
        Hour = 0,
        Minute = 0,
        Second = 0
    });

    var fc0 = new FilterColumn
    {
        ColumnId = 0U,
        HiddenButton = true,
        ShowButton = false
    };
    fc0.Append(filters);
    autoFilter.Append(fc0);

    var top10 = new Top10 { Top = true, Val = 5.0, FilterValue = 4.5 };
    var fc1 = new FilterColumn { ColumnId = 1U };
    fc1.Append(top10);
    autoFilter.Append(fc1);

    worksheetPart.Worksheet = new Worksheet(
        new SheetData(
            new Row(
                new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("Name") },
                new Cell { CellReference = "B1", DataType = CellValues.String, CellValue = new CellValue("Date") }
            ) { RowIndex = 1 }
        ),
        autoFilter
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
