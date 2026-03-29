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

// Add cell data
var sd = new SheetData();
var row = new Row { RowIndex = 1 };
row.Append(new Cell { CellReference = "A1", DataType = CellValues.Number, CellValue = new CellValue("50") });
sd.Append(row);
ws.Append(sd);

// Add dataBar CF with deep attributes
var cf1 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("A1:A10") }) };
var dbRule = new ConditionalFormattingRule { Type = ConditionalFormatValues.DataBar, Priority = 1 };
var dataBar = new DataBar { MinLength = 5, MaxLength = 90, ShowValue = false };
dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
dataBar.Append(new Color { Rgb = "FF638EC6" });
dbRule.Append(dataBar);
cf1.Append(dbRule);

// Add iconSet CF with deep attributes
var cf2 = new ConditionalFormatting { SequenceOfReferences = new ListValue<StringValue>(new[] { new StringValue("B1:B10") }) };
var isRule = new ConditionalFormattingRule { Type = ConditionalFormatValues.IconSet, Priority = 2 };
var iconSet = new IconSet { IconSetValue = IconSetValues.ThreeArrows, Reverse = true, ShowValue = false };
iconSet.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percent, Val = "0" });
iconSet.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percent, Val = "33" });
iconSet.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percent, Val = "67" });
isRule.Append(iconSet);
cf2.Append(isRule);

ws.Append(cf1);
ws.Append(cf2);

wsPart.Worksheet = ws;
wsPart.Worksheet.Save();
doc.Dispose();

Console.Error.WriteLine("SCENARIO_PASS");
