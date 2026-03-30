var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.Number,
        CellValue = new CellValue("50") })
);

// colorScale with theme colors
var cf1 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("A1:A10") }) };
var rule1 = new ConditionalFormattingRule { Type = ConditionalFormatValues.ColorScale, Priority = 1 };
var colorScale = new ColorScale();
colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
colorScale.Append(new Color { Theme = 4U, Tint = -0.25 });
colorScale.Append(new Color { Theme = 9U });
rule1.Append(colorScale);
cf1.Append(rule1);

// dataBar with indexed color
var cf2 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("B1:B10") }) };
var rule2 = new ConditionalFormattingRule { Type = ConditionalFormatValues.DataBar, Priority = 2 };
var dataBar = new DataBar();
dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
dataBar.Append(new Color { Indexed = 10U });
rule2.Append(dataBar);
cf2.Append(rule2);

worksheetPart.Worksheet = new Worksheet(sheetData, cf1, cf2);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
