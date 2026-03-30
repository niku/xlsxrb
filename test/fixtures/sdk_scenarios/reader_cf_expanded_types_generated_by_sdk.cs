var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("hello")) })
);

// aboveAverage rule (below average, equal)
var cf1 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("A1:A10") }) };
var rule1 = new ConditionalFormattingRule { Type = ConditionalFormatValues.AboveAverage, Priority = 1, AboveAverage = false, EqualAverage = true };
cf1.Append(rule1);

// top10 rule (bottom 5 percent)
var cf2 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("B1:B10") }) };
var rule2 = new ConditionalFormattingRule { Type = ConditionalFormatValues.Top10, Priority = 2, Rank = 5U, Percent = true, Bottom = true };
cf2.Append(rule2);

// duplicateValues rule
var cf3 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("C1:C10") }) };
var rule3 = new ConditionalFormattingRule { Type = ConditionalFormatValues.DuplicateValues, Priority = 3 };
cf3.Append(rule3);

// containsText rule
var cf4 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("D1:D10") }) };
var rule4 = new ConditionalFormattingRule { Type = ConditionalFormatValues.ContainsText, Priority = 4, Operator = ConditionalFormattingOperatorValues.ContainsText, Text = "hello" };
rule4.Append(new Formula("NOT(ISERROR(SEARCH(\"hello\",D1)))"));
cf4.Append(rule4);

// beginsWith rule
var cf5 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("E1:E10") }) };
var rule5 = new ConditionalFormattingRule { Type = ConditionalFormatValues.BeginsWith, Priority = 5, Operator = ConditionalFormattingOperatorValues.BeginsWith, Text = "foo" };
rule5.Append(new Formula("LEFT(E1,3)=\"foo\""));
cf5.Append(rule5);

// endsWith rule
var cf6 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("F1:F10") }) };
var rule6 = new ConditionalFormattingRule { Type = ConditionalFormatValues.EndsWith, Priority = 6, Operator = ConditionalFormattingOperatorValues.EndsWith, Text = "bar" };
rule6.Append(new Formula("RIGHT(F1,3)=\"bar\""));
cf6.Append(rule6);

worksheetPart.Worksheet = new Worksheet(sheetData, cf1, cf2, cf3, cf4, cf5, cf6);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
