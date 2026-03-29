// Creates an XLSX with a pivot table having col_fields, items, and populated cache.
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var wbPart = doc.AddWorkbookPart();
wbPart.Workbook = new Workbook();

var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");
var sd = new SheetData();
sd.Append(new Row(
    new Cell { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Category")) },
    new Cell { CellReference = "B1", DataType = CellValues.InlineString, InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Region")) },
    new Cell { CellReference = "C1", DataType = CellValues.InlineString, InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Amount")) }
));
wsPart.Worksheet = new Worksheet(sd);
wsPart.Worksheet.Save();

// Pivot cache definition
var cachePart = wbPart.AddNewPart<PivotTableCacheDefinitionPart>();
var cacheDef = new PivotCacheDefinition { RefreshOnLoad = true };
cacheDef.Append(new CacheSource(
    new WorksheetSource { Reference = "A1:C4", Sheet = "Sheet1" }
) { Type = SourceValues.Worksheet });
var cacheFields = new CacheFields { Count = 3 };
var cf0 = new CacheField { Name = "Category", NumberFormatId = 0 };
var si0 = new SharedItems { Count = 2 };
si0.Append(new StringItem { Val = "A" });
si0.Append(new StringItem { Val = "B" });
cf0.Append(si0);
cacheFields.Append(cf0);
var cf1 = new CacheField { Name = "Region", NumberFormatId = 0 };
var si1 = new SharedItems { Count = 2 };
si1.Append(new StringItem { Val = "East" });
si1.Append(new StringItem { Val = "West" });
cf1.Append(si1);
cacheFields.Append(cf1);
cacheFields.Append(new CacheField { Name = "Amount", NumberFormatId = 0, SharedItems = new SharedItems() });
cacheDef.Append(cacheFields);
cachePart.PivotCacheDefinition = cacheDef;
cachePart.PivotCacheDefinition.Save();

// Cache records
var recPart = cachePart.AddNewPart<PivotTableCacheRecordsPart>();
var records = new PivotCacheRecords { Count = 2 };
var r1 = new PivotCacheRecord();
r1.Append(new FieldItem { Val = 0 });
r1.Append(new FieldItem { Val = 0 });
r1.Append(new NumberItem { Val = 100 });
records.Append(r1);
var r2 = new PivotCacheRecord();
r2.Append(new FieldItem { Val = 1 });
r2.Append(new FieldItem { Val = 1 });
r2.Append(new NumberItem { Val = 200 });
records.Append(r2);
recPart.PivotCacheRecords = records;
recPart.PivotCacheRecords.Save();

// Pivot table
var ptPart = wsPart.AddNewPart<PivotTablePart>();
var ptDef = new PivotTableDefinition
{
    Name = "PivotDeep", CacheId = 0, DataCaption = "Values",
    DataOnRows = false,
    ApplyNumberFormats = false, ApplyBorderFormats = false,
    ApplyFontFormats = false, ApplyPatternFormats = false,
    ApplyAlignmentFormats = false, ApplyWidthHeightFormats = true
};
ptDef.Append(new Location { Reference = "E1:G5", FirstHeaderRow = 1, FirstDataRow = 1, FirstDataColumn = 1 });

var pivotFields = new PivotFields { Count = 3 };
// Field 0 = row
var pf0 = new PivotField { Axis = PivotTableAxisValues.AxisRow, ShowAll = false };
var items0 = new Items { Count = 3 };
items0.Append(new Item { Index = 0 });
items0.Append(new Item { Index = 1 });
items0.Append(new Item { ItemType = ItemValues.Default });
pf0.Append(items0);
pivotFields.Append(pf0);
// Field 1 = col
var pf1 = new PivotField { Axis = PivotTableAxisValues.AxisColumn, ShowAll = false };
var items1 = new Items { Count = 3 };
items1.Append(new Item { Index = 0 });
items1.Append(new Item { Index = 1 });
items1.Append(new Item { ItemType = ItemValues.Default });
pf1.Append(items1);
pivotFields.Append(pf1);
// Field 2 = data
pivotFields.Append(new PivotField { DataField = true, ShowAll = false });

ptDef.Append(pivotFields);
ptDef.Append(new RowFields(new Field { Index = 0 }) { Count = 1 });
ptDef.Append(new ColumnFields(new Field { Index = 1 }) { Count = 1 });
ptDef.Append(new DataFields(new DataField { Name = "Sum of Amount", Field = 2, Subtotal = DataConsolidateFunctionValues.Sum }) { Count = 1 });

ptPart.PivotTableDefinition = ptDef;
ptPart.PivotTableDefinition.Save();

// Link pivot table to cache
ptPart.AddPart(cachePart, "rId1");

// Workbook pivot cache definitions
var cacheRid = wbPart.GetIdOfPart(cachePart);
wbPart.Workbook.Append(new PivotCaches(new PivotCache { CacheId = 0, Id = cacheRid }));

var sheetsEl = wbPart.Workbook.AppendChild(new Sheets());
sheetsEl.Append(new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" });
wbPart.Workbook.Save();
doc.Dispose();

Console.Error.WriteLine("SCENARIO_PASS");
