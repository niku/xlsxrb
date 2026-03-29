var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(
        new Cell { CellReference = "A1", DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text("Name")) },
        new Cell { CellReference = "B1", DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text("Score")) }
    ),
    new Row(
        new Cell { CellReference = "A2", DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text("Alice")) },
        new Cell { CellReference = "B2", DataType = CellValues.Number,
            CellValue = new CellValue("95") }
    )
);

var tableParts = new TableParts { Count = 1 };
worksheetPart.Worksheet = new Worksheet(sheetData, tableParts);

// Add table definition part
var tableDefPart = worksheetPart.AddNewPart<TableDefinitionPart>();
var table = new Table
{
    Id = 1, Name = "SdkTable", DisplayName = "SdkTable",
    Reference = "A1:B2", TotalsRowShown = false
};
table.AutoFilter = new AutoFilter { Reference = "A1:B2" };
var tableCols = new TableColumns { Count = 2 };
tableCols.Append(new TableColumn { Id = 1, Name = "Name" });
tableCols.Append(new TableColumn { Id = 2, Name = "Score" });
table.Append(tableCols);
table.Append(new TableStyleInfo
{
    Name = "TableStyleMedium2",
    ShowFirstColumn = false, ShowLastColumn = false,
    ShowRowStripes = true, ShowColumnStripes = false
});
tableDefPart.Table = table;
tableDefPart.Table.Save();

// Wire up tablePart reference
var tpRid = worksheetPart.GetIdOfPart(tableDefPart);
tableParts.Append(new TablePart { Id = tpRid });
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
