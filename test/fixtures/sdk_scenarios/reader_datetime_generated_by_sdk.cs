var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
    stylesPart.Stylesheet = new Stylesheet(
        new NumberingFormats(
            new NumberingFormat { NumberFormatId = 164, FormatCode = "yyyy-mm-dd hh:mm:ss" }
        ) { Count = 1 },
        new Fonts(new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" })) { Count = 1 },
        new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
        ) { Count = 2 },
        new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())) { Count = 1 },
        new CellStyleFormats(new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }) { Count = 1 },
        new CellFormats(
            new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 },
            new CellFormat { NumberFormatId = 164, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0, ApplyNumberFormat = true }
        ) { Count = 2 },
        new CellStyles(new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 }) { Count = 1 }
    );
    stylesPart.Stylesheet.Save();

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    // 2024-06-15 10:15:30 => serial 45458 + (10*3600+15*60+30)/86400 = 45458.427430555...
    var sheetData = new SheetData(
        new Row(
            new Cell { CellReference = "A1", CellValue = new CellValue("45458.427430555556"), StyleIndex = 1 }
        )
    );
    worksheetPart.Worksheet = new Worksheet(sheetData);

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
