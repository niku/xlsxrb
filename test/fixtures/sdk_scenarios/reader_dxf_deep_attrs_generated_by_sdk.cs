var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

// Styles with DXF including numFmt, alignment, protection
var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
var stylesheet = new Stylesheet();

var fonts = new Fonts(new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" }));
fonts.Count = 1;
stylesheet.Append(fonts);

var fills = new Fills(
    new Fill(new PatternFill { PatternType = PatternValues.None }),
    new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
);
fills.Count = 2;
stylesheet.Append(fills);

var borders = new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()));
borders.Count = 1;
stylesheet.Append(borders);

var cellStyleFormats = new CellStyleFormats(new CellFormat { FontId = 0, FillId = 0, BorderId = 0 });
cellStyleFormats.Count = 1;
stylesheet.Append(cellStyleFormats);

var cellFormats = new CellFormats(new CellFormat { FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 });
cellFormats.Count = 1;
stylesheet.Append(cellFormats);

var cellStyles = new CellStyles(new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 });
cellStyles.Count = 1;
stylesheet.Append(cellStyles);

// DXF with font, numFmt, alignment, protection
var dxfs = new DifferentialFormats();
var dxf = new DifferentialFormat();
dxf.Append(new Font(new Bold(), new Color { Rgb = "FFFF0000" }));
dxf.Append(new NumberingFormat { NumberFormatId = 164, FormatCode = "#,##0.00" });
dxf.Append(new Alignment { Horizontal = HorizontalAlignmentValues.Center, WrapText = true });
dxf.Append(new Protection { Locked = false, Hidden = true });
dxfs.Append(dxf);
dxfs.Count = 1;
stylesheet.Append(dxfs);

stylesPart.Stylesheet = stylesheet;
stylesPart.Stylesheet.Save();

// Worksheet
var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("hello")) })
);
worksheetPart.Worksheet = new Worksheet(sheetData);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
