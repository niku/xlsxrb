var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;

var allRules = worksheet.Elements<ConditionalFormatting>()
    .SelectMany(cf => cf.Elements<ConditionalFormattingRule>()).ToList();
if (allRules.Count != 6)
    throw new Exception($"SCENARIO_FAIL: Expected 6 cfRules, got {allRules.Count}");

// aboveAverage rule
var r0 = allRules[0];
if (r0.Type?.Value != ConditionalFormatValues.AboveAverage)
    throw new Exception($"SCENARIO_FAIL: Rule0 type expected AboveAverage, got {r0.Type?.Value}");
if (r0.AboveAverage?.Value != false)
    throw new Exception($"SCENARIO_FAIL: Rule0 AboveAverage expected false, got {r0.AboveAverage?.Value}");
if (r0.EqualAverage?.Value != true)
    throw new Exception($"SCENARIO_FAIL: Rule0 EqualAverage expected true, got {r0.EqualAverage?.Value}");

// top10 rule
var r1 = allRules[1];
if (r1.Type?.Value != ConditionalFormatValues.Top10)
    throw new Exception($"SCENARIO_FAIL: Rule1 type expected Top10, got {r1.Type?.Value}");
if (r1.Rank?.Value != 5U)
    throw new Exception($"SCENARIO_FAIL: Rule1 Rank expected 5, got {r1.Rank?.Value}");
if (r1.Percent?.Value != true)
    throw new Exception($"SCENARIO_FAIL: Rule1 Percent expected true, got {r1.Percent?.Value}");
if (r1.Bottom?.Value != true)
    throw new Exception($"SCENARIO_FAIL: Rule1 Bottom expected true, got {r1.Bottom?.Value}");

// duplicateValues rule
var r2 = allRules[2];
if (r2.Type?.Value != ConditionalFormatValues.DuplicateValues)
    throw new Exception($"SCENARIO_FAIL: Rule2 type expected DuplicateValues, got {r2.Type?.Value}");

// containsText rule
var r3 = allRules[3];
if (r3.Type?.Value != ConditionalFormatValues.ContainsText)
    throw new Exception($"SCENARIO_FAIL: Rule3 type expected ContainsText, got {r3.Type?.Value}");
if (r3.Text?.Value != "hello")
    throw new Exception($"SCENARIO_FAIL: Rule3 Text expected 'hello', got '{r3.Text?.Value}'");
var formulas3 = r3.Elements<Formula>().ToList();
if (formulas3.Count != 1)
    throw new Exception($"SCENARIO_FAIL: Rule3 expected 1 formula, got {formulas3.Count}");

// beginsWith rule
var r4 = allRules[4];
if (r4.Type?.Value != ConditionalFormatValues.BeginsWith)
    throw new Exception($"SCENARIO_FAIL: Rule4 type expected BeginsWith, got {r4.Type?.Value}");
if (r4.Text?.Value != "foo")
    throw new Exception($"SCENARIO_FAIL: Rule4 Text expected 'foo', got '{r4.Text?.Value}'");

// endsWith rule
var r5 = allRules[5];
if (r5.Type?.Value != ConditionalFormatValues.EndsWith)
    throw new Exception($"SCENARIO_FAIL: Rule5 type expected EndsWith, got {r5.Type?.Value}");
if (r5.Text?.Value != "bar")
    throw new Exception($"SCENARIO_FAIL: Rule5 Text expected 'bar', got '{r5.Text?.Value}'");

doc.Dispose();
