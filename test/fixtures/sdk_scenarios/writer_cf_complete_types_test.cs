var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;

var allRules = worksheet.Elements<ConditionalFormatting>()
    .SelectMany(cf => cf.Elements<ConditionalFormattingRule>()).ToList();
if (allRules.Count != 6)
    throw new Exception($"SCENARIO_FAIL: Expected 6 cfRules, got {allRules.Count}");

// expression
var r0 = allRules[0];
if (r0.Type?.Value != ConditionalFormatValues.Expression)
    throw new Exception($"SCENARIO_FAIL: Rule0 type expected Expression, got {r0.Type?.Value}");
var formulas0 = r0.Elements<Formula>().ToList();
if (formulas0.Count != 1)
    throw new Exception($"SCENARIO_FAIL: Rule0 expected 1 formula, got {formulas0.Count}");

// uniqueValues
var r1 = allRules[1];
if (r1.Type?.Value != ConditionalFormatValues.UniqueValues)
    throw new Exception($"SCENARIO_FAIL: Rule1 type expected UniqueValues, got {r1.Type?.Value}");

// notContainsText
var r2 = allRules[2];
if (r2.Type?.Value != ConditionalFormatValues.NotContainsText)
    throw new Exception($"SCENARIO_FAIL: Rule2 type expected NotContainsText, got {r2.Type?.Value}");
if (r2.Text?.Value != "bad")
    throw new Exception($"SCENARIO_FAIL: Rule2 Text expected 'bad', got '{r2.Text?.Value}'");

// containsBlanks
var r3 = allRules[3];
if (r3.Type?.Value != ConditionalFormatValues.ContainsBlanks)
    throw new Exception($"SCENARIO_FAIL: Rule3 type expected ContainsBlanks, got {r3.Type?.Value}");

// notContainsBlanks
var r4 = allRules[4];
if (r4.Type?.Value != ConditionalFormatValues.NotContainsBlanks)
    throw new Exception($"SCENARIO_FAIL: Rule4 type expected NotContainsBlanks, got {r4.Type?.Value}");

// timePeriod
var r5 = allRules[5];
if (r5.Type?.Value != ConditionalFormatValues.TimePeriod)
    throw new Exception($"SCENARIO_FAIL: Rule5 type expected TimePeriod, got {r5.Type?.Value}");
if (r5.TimePeriod?.Value != TimePeriodValues.LastWeek)
    throw new Exception($"SCENARIO_FAIL: Rule5 timePeriod expected LastWeek, got {r5.TimePeriod?.Value}");

doc.Dispose();
