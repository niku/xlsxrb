var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;

var dvs = worksheet.GetFirstChild<DataValidations>();
if (dvs == null)
    throw new Exception("SCENARIO_FAIL: DataValidations is null");

var dvList = dvs.Elements<DataValidation>().ToList();
if (dvList.Count != 2)
    throw new Exception($"SCENARIO_FAIL: Expected 2 DataValidation, got {dvList.Count}");

var dv0 = dvList[0];
if (dv0.SequenceOfReferences?.InnerText != "A1:A100")
    throw new Exception($"SCENARIO_FAIL: dv0 sqref expected A1:A100, got {dv0.SequenceOfReferences?.InnerText}");
if (dv0.Type?.Value != DataValidationValues.Whole)
    throw new Exception($"SCENARIO_FAIL: dv0 type expected Whole, got {dv0.Type?.Value}");
if (dv0.Operator?.Value != DataValidationOperatorValues.Between)
    throw new Exception($"SCENARIO_FAIL: dv0 operator expected Between, got {dv0.Operator?.Value}");
if (dv0.ShowErrorMessage?.Value != true)
    throw new Exception($"SCENARIO_FAIL: dv0 showErrorMessage expected true, got {dv0.ShowErrorMessage?.Value}");
if (dv0.Error?.Value != "Must be 1-100")
    throw new Exception($"SCENARIO_FAIL: dv0 error expected 'Must be 1-100', got {dv0.Error?.Value}");
var f1 = dv0.Formula1?.Text;
if (f1 != "1")
    throw new Exception($"SCENARIO_FAIL: dv0 formula1 expected 1, got {f1}");
var f2 = dv0.Formula2?.Text;
if (f2 != "100")
    throw new Exception($"SCENARIO_FAIL: dv0 formula2 expected 100, got {f2}");

var dv1 = dvList[1];
if (dv1.SequenceOfReferences?.InnerText != "B1:B100")
    throw new Exception($"SCENARIO_FAIL: dv1 sqref expected B1:B100, got {dv1.SequenceOfReferences?.InnerText}");
if (dv1.Type?.Value != DataValidationValues.List)
    throw new Exception($"SCENARIO_FAIL: dv1 type expected List, got {dv1.Type?.Value}");
if (dv1.ShowInputMessage?.Value != true)
    throw new Exception($"SCENARIO_FAIL: dv1 showInputMessage expected true, got {dv1.ShowInputMessage?.Value}");

doc.Dispose();
