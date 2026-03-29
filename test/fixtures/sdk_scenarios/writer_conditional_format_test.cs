var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;

var cfList = worksheet.Elements<ConditionalFormatting>().ToList();
if (cfList.Count == 0)
    throw new Exception("SCENARIO_FAIL: No ConditionalFormatting elements found");

// Collect all cfRules across all conditionalFormatting elements
var allRules = cfList.SelectMany(cf => cf.Elements<ConditionalFormattingRule>()).ToList();
if (allRules.Count != 4)
    throw new Exception($"SCENARIO_FAIL: Expected 4 cfRules total, got {allRules.Count}");

// cellIs rule
var r0 = allRules[0];
if (r0.Type?.Value != ConditionalFormatValues.CellIs)
    throw new Exception($"SCENARIO_FAIL: Rule0 type expected CellIs, got {r0.Type?.Value}");
if (r0.Operator?.Value != ConditionalFormattingOperatorValues.GreaterThan)
    throw new Exception($"SCENARIO_FAIL: Rule0 operator expected GreaterThan, got {r0.Operator?.Value}");
var formulas = r0.Elements<Formula>().ToList();
if (formulas.Count != 1 || formulas[0].Text != "100")
    throw new Exception($"SCENARIO_FAIL: Rule0 formula expected '100', got {formulas.FirstOrDefault()?.Text}");

// colorScale rule
var r1 = allRules[1];
if (r1.Type?.Value != ConditionalFormatValues.ColorScale)
    throw new Exception($"SCENARIO_FAIL: Rule1 type expected ColorScale, got {r1.Type?.Value}");
var cs = r1.GetFirstChild<ColorScale>();
if (cs == null)
    throw new Exception("SCENARIO_FAIL: ColorScale element is null");
var cfvos = cs.Elements<ConditionalFormatValueObject>().ToList();
if (cfvos.Count != 2)
    throw new Exception($"SCENARIO_FAIL: ColorScale expected 2 cfvo, got {cfvos.Count}");

// dataBar rule
var r2 = allRules[2];
if (r2.Type?.Value != ConditionalFormatValues.DataBar)
    throw new Exception($"SCENARIO_FAIL: Rule2 type expected DataBar, got {r2.Type?.Value}");
var db = r2.GetFirstChild<DataBar>();
if (db == null)
    throw new Exception("SCENARIO_FAIL: DataBar element is null");

// iconSet rule
var r3 = allRules[3];
if (r3.Type?.Value != ConditionalFormatValues.IconSet)
    throw new Exception($"SCENARIO_FAIL: Rule3 type expected IconSet, got {r3.Type?.Value}");
var iconSet = r3.GetFirstChild<IconSet>();
if (iconSet == null)
    throw new Exception("SCENARIO_FAIL: IconSet element is null");

doc.Dispose();
