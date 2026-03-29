var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("Value")) })
);

var dvs = new DataValidations();
var dv0 = new DataValidation
{
    Type = DataValidationValues.Whole,
    Operator = DataValidationOperatorValues.GreaterThanOrEqual,
    ShowErrorMessage = true,
    Error = "Must be >= 10",
    ErrorTitle = "Validation Error",
    SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("A2:A50") })
};
dv0.Formula1 = new Formula1("10");
dvs.Append(dv0);

var dv1 = new DataValidation
{
    Type = DataValidationValues.List,
    ShowInputMessage = true,
    Prompt = "Pick a color",
    PromptTitle = "Color",
    SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("B2:B50") })
};
dv1.Formula1 = new Formula1("\"Red,Green,Blue\"");
dvs.Append(dv1);
dvs.Count = 2;

worksheetPart.Worksheet = new Worksheet(sheetData, dvs);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
