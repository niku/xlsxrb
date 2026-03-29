using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var wbPart = doc.AddWorkbookPart();
wbPart.Workbook = new Workbook(new Sheets(
    new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }
));
var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");
var ws = new Worksheet();

var sd = new SheetData();
var row = new Row { RowIndex = 1 };
row.Append(new Cell { CellReference = "A1", DataType = CellValues.Number, CellValue = new CellValue("50") });
sd.Append(row);
ws.Append(sd);

var dvs = new DataValidations();
var dv = new DataValidation
{
    SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("A1:A100") }),
    Type = DataValidationValues.Whole,
    Operator = DataValidationOperatorValues.Between,
    AllowBlank = true,
    ErrorStyle = DataValidationErrorStyleValues.Warning,
    ErrorTitle = "Bad Value",
    Error = "Please enter 1-100",
    ShowErrorMessage = true,
    PromptTitle = "Input Needed",
    Prompt = "Enter a number",
    ShowInputMessage = true
};
dv.Append(new Formula1("1"));
dv.Append(new Formula2("100"));
dvs.Append(dv);
ws.Append(dvs);

wsPart.Worksheet = ws;
wsPart.Worksheet.Save();
doc.Dispose();

Console.Error.WriteLine("SCENARIO_PASS");
