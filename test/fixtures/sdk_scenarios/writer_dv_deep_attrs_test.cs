using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).Take(10).ToList();
    if (errors.Count > 0)
    {
        foreach (var e in errors)
            Console.Error.WriteLine("Validation: " + e.Description);
        throw new Exception("SCENARIO_FAIL: validation errors: " + errors.Count);
    }

    var wbPart = document.WorkbookPart;
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var dvs = wsPart.Worksheet.Elements<DataValidations>().FirstOrDefault();

    if (dvs == null)
        throw new Exception("SCENARIO_FAIL: no dataValidations element found");

    var dvList = dvs.Elements<DataValidation>().ToList();
    if (dvList.Count < 1)
        throw new Exception("SCENARIO_FAIL: no dataValidation elements found");

    var dv = dvList[0];

    if (dv.AllowBlank == null || dv.AllowBlank.Value != true)
        throw new Exception("SCENARIO_FAIL: allowBlank expected true");
    if (dv.ErrorStyle == null || dv.ErrorStyle.Value != DataValidationErrorStyleValues.Warning)
        throw new Exception("SCENARIO_FAIL: errorStyle expected warning");
    if (dv.ErrorTitle == null || dv.ErrorTitle.Value != "Bad Value")
        throw new Exception("SCENARIO_FAIL: errorTitle expected 'Bad Value', got " + dv.ErrorTitle);
    if (dv.Error == null || dv.Error.Value != "Please enter 1-100")
        throw new Exception("SCENARIO_FAIL: error expected 'Please enter 1-100'");
    if (dv.PromptTitle == null || dv.PromptTitle.Value != "Input Needed")
        throw new Exception("SCENARIO_FAIL: promptTitle expected 'Input Needed'");
    if (dv.Prompt == null || dv.Prompt.Value != "Enter a number")
        throw new Exception("SCENARIO_FAIL: prompt expected 'Enter a number'");
    if (dv.ShowInputMessage == null || dv.ShowInputMessage.Value != true)
        throw new Exception("SCENARIO_FAIL: showInputMessage expected true");
    if (dv.ShowErrorMessage == null || dv.ShowErrorMessage.Value != true)
        throw new Exception("SCENARIO_FAIL: showErrorMessage expected true");

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
