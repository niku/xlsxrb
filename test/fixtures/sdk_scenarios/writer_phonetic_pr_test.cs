var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;
var phoneticPr = worksheet.GetFirstChild<PhoneticProperties>();

if (phoneticPr == null)
    throw new Exception("SCENARIO_FAIL: PhoneticProperties is null");

if (phoneticPr.FontId?.Value != (uint)1)
    throw new Exception($"SCENARIO_FAIL: FontId expected 1, got {phoneticPr.FontId?.Value}");
if (phoneticPr.Type?.Value != PhoneticValues.Hiragana)
    throw new Exception($"SCENARIO_FAIL: Type expected Hiragana, got {phoneticPr.Type?.Value}");
if (phoneticPr.Alignment?.Value != PhoneticAlignmentValues.Center)
    throw new Exception($"SCENARIO_FAIL: Alignment expected Center, got {phoneticPr.Alignment?.Value}");

Console.WriteLine("SCENARIO_PASS");
