var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

// Create styles
var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
var stylesheet = new Stylesheet();

// Fonts
var fonts = new Fonts(
    new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" }),
    new Font(new Bold(), new FontSize { Val = 16 }, new FontName { Val = "Times New Roman" }, new Color { Rgb = "FF0000FF" })
);
fonts.Count = 2;
stylesheet.Fonts = fonts;

// Fills
var fills = new Fills(
    new Fill(new PatternFill { PatternType = PatternValues.None }),
    new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
    new Fill(new PatternFill(new ForegroundColor { Rgb = "FFFF0000" }) { PatternType = PatternValues.Solid })
);
fills.Count = 3;
stylesheet.Fills = fills;

// Borders
var borders = new Borders(
    new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()),
    new Border(
        new LeftBorder(new Color { Rgb = "FF000000" }) { Style = BorderStyleValues.Medium },
        new RightBorder { Style = BorderStyleValues.Medium },
        new TopBorder { Style = BorderStyleValues.Medium },
        new BottomBorder { Style = BorderStyleValues.Medium },
        new DiagonalBorder()
    )
);
borders.Count = 2;
stylesheet.Borders = borders;

// CellStyleXfs
stylesheet.CellStyleFormats = new CellStyleFormats(
    new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 }
) { Count = 1 };

// CellXfs
stylesheet.CellFormats = new CellFormats(
    new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 },
    new CellFormat { NumberFormatId = 0, FontId = 1, FillId = 2, BorderId = 1, FormatId = 0,
        ApplyFont = true, ApplyFill = true, ApplyBorder = true }
) { Count = 2 };

// DXFs
stylesheet.DifferentialFormats = new DifferentialFormats(
    new DifferentialFormat(
        new Font(new Bold(), new Color { Rgb = "FF00FF00" }),
        new Fill(new PatternFill(new ForegroundColor { Rgb = "FFFFFF00" }) { PatternType = PatternValues.Solid })
    )
) { Count = 1 };

stylesheet.Save(stylesPart);

// Worksheet
var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", StyleIndex = 1, DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("Styled")) })
);
worksheetPart.Worksheet = new Worksheet(sheetData);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
