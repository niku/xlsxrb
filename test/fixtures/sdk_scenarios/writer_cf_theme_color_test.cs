var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;

var allRules = worksheet.Elements<ConditionalFormatting>()
    .SelectMany(cf => cf.Elements<ConditionalFormattingRule>()).ToList();
if (allRules.Count != 2)
    throw new Exception($"SCENARIO_FAIL: Expected 2 cfRules, got {allRules.Count}");

// colorScale with theme colors
var r0 = allRules[0];
if (r0.Type?.Value != ConditionalFormatValues.ColorScale)
    throw new Exception($"SCENARIO_FAIL: Rule0 type expected ColorScale, got {r0.Type?.Value}");
var cs = r0.GetFirstChild<ColorScale>();
if (cs == null)
    throw new Exception("SCENARIO_FAIL: ColorScale is null");
var colors = cs.Elements<Color>().ToList();
if (colors.Count != 2)
    throw new Exception($"SCENARIO_FAIL: Expected 2 colors, got {colors.Count}");
if (colors[0].Theme?.Value != 4U)
    throw new Exception($"SCENARIO_FAIL: Color0 theme expected 4, got {colors[0].Theme?.Value}");
if (colors[0].Tint?.Value == null || Math.Abs(colors[0].Tint.Value - (-0.25)) > 0.01)
    throw new Exception($"SCENARIO_FAIL: Color0 tint expected -0.25, got {colors[0].Tint?.Value}");
if (colors[1].Theme?.Value != 9U)
    throw new Exception($"SCENARIO_FAIL: Color1 theme expected 9, got {colors[1].Theme?.Value}");

// dataBar with indexed color
var r1 = allRules[1];
if (r1.Type?.Value != ConditionalFormatValues.DataBar)
    throw new Exception($"SCENARIO_FAIL: Rule1 type expected DataBar, got {r1.Type?.Value}");
var db = r1.GetFirstChild<DataBar>();
if (db == null)
    throw new Exception("SCENARIO_FAIL: DataBar is null");
var dbColor = db.GetFirstChild<Color>();
if (dbColor == null || dbColor.Indexed?.Value != 10U)
    throw new Exception($"SCENARIO_FAIL: DataBar color indexed expected 10, got {dbColor?.Indexed?.Value}");

doc.Dispose();
