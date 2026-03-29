var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("Val")) })
);

// cellIs rule
var cf1 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("A1:A10") }) };
var rule1 = new ConditionalFormattingRule { Type = ConditionalFormatValues.CellIs, Operator = ConditionalFormattingOperatorValues.LessThan, Priority = 1 };
rule1.Append(new Formula("50"));
cf1.Append(rule1);

// colorScale rule
var cf2 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("B1:B10") }) };
var rule2 = new ConditionalFormattingRule { Type = ConditionalFormatValues.ColorScale, Priority = 2 };
var colorScale = new ColorScale();
colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
colorScale.Append(new Color { Rgb = "FF00FF00" });
colorScale.Append(new Color { Rgb = "FFFF0000" });
rule2.Append(colorScale);
cf2.Append(rule2);

// dataBar rule
var cf3 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("C1:C10") }) };
var rule3 = new ConditionalFormattingRule { Type = ConditionalFormatValues.DataBar, Priority = 3 };
var dataBar = new DataBar();
dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
dataBar.Append(new Color { Rgb = "FF63C384" });
rule3.Append(dataBar);
cf3.Append(rule3);

worksheetPart.Worksheet = new Worksheet(sheetData, cf1, cf2, cf3);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
