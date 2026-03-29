// Creates an XLSX with a pivot table.
var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

// Source data sheet.
var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(
        new Cell { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Category")) },
        new Cell { CellReference = "B1", DataType = CellValues.InlineString, InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Amount")) }
    ),
    new Row(
        new Cell { CellReference = "A2", DataType = CellValues.InlineString, InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("A")) },
        new Cell { CellReference = "B2", CellValue = new CellValue("100") }
    ),
    new Row(
        new Cell { CellReference = "A3", DataType = CellValues.InlineString, InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("B")) },
        new Cell { CellReference = "B3", CellValue = new CellValue("200") }
    )
);
worksheetPart.Worksheet = new Worksheet(sheetData);
worksheetPart.Worksheet.Save();

// Pivot sheet.
var pivotSheetPart = workbookPart.AddNewPart<WorksheetPart>();
pivotSheetPart.Worksheet = new Worksheet(new SheetData());
pivotSheetPart.Worksheet.Save();

// Create pivot cache.
var pivotCachePart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
var cacheDef = new PivotCacheDefinition { RecordCount = 2 };
var cacheSource = new CacheSource { Type = SourceValues.Worksheet };
cacheSource.Append(new WorksheetSource { Reference = "A1:B3", Sheet = "Sheet1" });
cacheDef.Append(cacheSource);
var cacheFields = new CacheFields();
var cf1 = new CacheField { Name = "Category", NumberFormatId = 0 };
cf1.Append(new SharedItems());
cacheFields.Append(cf1);
var cf2 = new CacheField { Name = "Amount", NumberFormatId = 0 };
cf2.Append(new SharedItems { ContainsSemiMixedTypes = false, ContainsString = false, ContainsNumber = true });
cacheFields.Append(cf2);
cacheDef.Append(cacheFields);
pivotCachePart.PivotCacheDefinition = cacheDef;
pivotCachePart.PivotCacheDefinition.Save();

// Create pivot cache records.
var recordsPart = pivotCachePart.AddNewPart<PivotTableCacheRecordsPart>();
recordsPart.PivotCacheRecords = new PivotCacheRecords { Count = 0 };
recordsPart.PivotCacheRecords.Save();

// Create pivot table.
var pivotTablePart = pivotSheetPart.AddNewPart<PivotTablePart>();
pivotTablePart.AddPart(pivotCachePart, "rId1");

var cacheId = 1U;
workbookPart.Workbook.Append(new PivotCaches(new PivotCache { CacheId = cacheId, Id = workbookPart.GetIdOfPart(pivotCachePart) }));

var ptDef = new PivotTableDefinition { Name = "TestPivot", CacheId = cacheId };
ptDef.Append(new Location { Reference = "A1:B3", FirstHeaderRow = 1, FirstDataRow = 1, FirstDataColumn = 1 });
var pvtFields = new PivotFields();
pvtFields.Append(new PivotField { Axis = PivotTableAxisValues.AxisRow, ShowAll = false });
pvtFields.Append(new PivotField { DataField = true, ShowAll = false });
pvtFields.Count = 2;
ptDef.Append(pvtFields);

ptDef.Append(new RowFields(new Field { Index = 0 }) { Count = 1 });
ptDef.Append(new DataFields(new DataField { Name = "Sum of Amount", Field = 1, Subtotal = DataConsolidateFunctionValues.Sum }) { Count = 1 });

pivotTablePart.PivotTableDefinition = ptDef;
pivotTablePart.PivotTableDefinition.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(pivotSheetPart), SheetId = 2, Name = "PivotSheet" });
workbookPart.Workbook.Save();
doc.Dispose();
