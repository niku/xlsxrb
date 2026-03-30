var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("hello")) })
);

// expression
var cf1 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("A1:A10") }) };
var rule1 = new ConditionalFormattingRule { Type = ConditionalFormatValues.Expression, Priority = 1 };
rule1.Append(new Formula("MOD(ROW(),2)=0"));
cf1.Append(rule1);

// uniqueValues
var cf2 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("B1:B10") }) };
var rule2 = new ConditionalFormattingRule { Type = ConditionalFormatValues.UniqueValues, Priority = 2 };
cf2.Append(rule2);

// notContainsText
var cf3 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("C1:C10") }) };
var rule3 = new ConditionalFormattingRule { Type = ConditionalFormatValues.NotContainsText, Priority = 3, Operator = ConditionalFormattingOperatorValues.NotContains, Text = "bad" };
rule3.Append(new Formula("ISERROR(SEARCH(\"bad\",C1))"));
cf3.Append(rule3);

// containsBlanks
var cf4 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("D1:D10") }) };
var rule4 = new ConditionalFormattingRule { Type = ConditionalFormatValues.ContainsBlanks, Priority = 4 };
rule4.Append(new Formula("LEN(TRIM(D1))=0"));
cf4.Append(rule4);

// notContainsBlanks
var cf5 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("E1:E10") }) };
var rule5 = new ConditionalFormattingRule { Type = ConditionalFormatValues.NotContainsBlanks, Priority = 5 };
rule5.Append(new Formula("LEN(TRIM(E1))>0"));
cf5.Append(rule5);

// timePeriod
var cf6 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("F1:F10") }) };
var rule6 = new ConditionalFormattingRule { Type = ConditionalFormatValues.TimePeriod, Priority = 6, TimePeriod = TimePeriodValues.LastWeek };
rule6.Append(new Formula("AND(TODAY()-7<=F1,F1<=TODAY())"));
cf6.Append(rule6);

worksheetPart.Worksheet = new Worksheet(sheetData, cf1, cf2, cf3, cf4, cf5, cf6);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
