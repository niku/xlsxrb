var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("Name")) }),
    new Row(new Cell { CellReference = "A2", DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("Alice")) })
);
var autoFilter = new AutoFilter { Reference = "A1:B10" };
autoFilter.Append(new FilterColumn(
    new Filters(new Filter { Val = "Alice" }, new Filter { Val = "Bob" })
) { ColumnId = 0 });
autoFilter.Append(new FilterColumn(
    new CustomFilters(new CustomFilter { Operator = FilterOperatorValues.GreaterThan, Val = "50" })
) { ColumnId = 1 });

var sortState = new SortState { Reference = "A2:B10" };
sortState.Append(new SortCondition { Reference = "A2:A10" });
sortState.Append(new SortCondition { Reference = "B2:B10", Descending = true });

worksheetPart.Worksheet = new Worksheet(sheetData, autoFilter, sortState);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
