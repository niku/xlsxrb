var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var stylesPart = workbookPart.WorkbookStylesPart;
var stylesheet = stylesPart.Stylesheet;

// Check fonts
var fonts = stylesheet.Fonts;
if (fonts.Count < 2)
    throw new Exception($"SCENARIO_FAIL: Expected at least 2 fonts, got {fonts.Count}");
var font1 = fonts.Elements<Font>().ElementAt(1);
if (font1.GetFirstChild<Bold>() == null)
    throw new Exception("SCENARIO_FAIL: Font 1 should be bold");
var fontSz = font1.GetFirstChild<FontSize>();
if (fontSz?.Val?.Value != 14.0)
    throw new Exception($"SCENARIO_FAIL: Font 1 size expected 14, got {fontSz?.Val?.Value}");

// Check fills
var fills = stylesheet.Fills;
if (fills.Count < 3)
    throw new Exception($"SCENARIO_FAIL: Expected at least 3 fills, got {fills.Count}");

// Check borders
var borders = stylesheet.Borders;
if (borders.Count < 2)
    throw new Exception($"SCENARIO_FAIL: Expected at least 2 borders, got {borders.Count}");

// Check cellXfs
var cellXfs = stylesheet.CellFormats;
if (cellXfs.Count < 2)
    throw new Exception($"SCENARIO_FAIL: Expected at least 2 xf entries, got {cellXfs.Count}");
var xf1 = cellXfs.Elements<CellFormat>().ElementAt(1);
if (xf1.FontId?.Value != 1)
    throw new Exception($"SCENARIO_FAIL: xf1 fontId expected 1, got {xf1.FontId?.Value}");

// Check dxfs
var dxfs = stylesheet.DifferentialFormats;
if (dxfs == null || dxfs.Count < 1)
    throw new Exception($"SCENARIO_FAIL: Expected at least 1 dxf, got {dxfs?.Count}");
var dxf0 = dxfs.Elements<DifferentialFormat>().First();
if (dxf0.GetFirstChild<Font>()?.GetFirstChild<Bold>() == null)
    throw new Exception("SCENARIO_FAIL: dxf0 font should be bold");

// Check cell A1 has style index
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var cells = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().First().Elements<Cell>().ToList();
var cellA1 = cells.First(c => c.CellReference?.Value == "A1");
if (cellA1.StyleIndex?.Value != 1)
    throw new Exception($"SCENARIO_FAIL: A1 styleIndex expected 1, got {cellA1.StyleIndex?.Value}");

doc.Dispose();
