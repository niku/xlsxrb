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
var sd = new SheetData();
var row1 = new Row { RowIndex = 1 };
row1.Append(new Cell { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Item")) });
row1.Append(new Cell { CellReference = "B1", DataType = CellValues.InlineString, InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Price")) });
row1.Append(new Cell { CellReference = "C1", DataType = CellValues.InlineString, InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Tax")) });
sd.Append(row1);
ws.Append(sd);

// Add tableParts reference
var tableParts = new TableParts { Count = 1 };
tableParts.Append(new TablePart { Id = "rId2" });
ws.Append(tableParts);

wsPart.Worksheet = ws;
wsPart.Worksheet.Save();

// Create table part
var tablePart = wsPart.AddNewPart<TableDefinitionPart>("rId2");
var table = new Table
{
    Id = 1, Name = "SalesTable", DisplayName = "SalesTable", Reference = "A1:C5",
    TotalsRowCount = 1
};
table.Append(new AutoFilter { Reference = "A1:C5" });

var tableCols = new TableColumns { Count = 3 };
tableCols.Append(new TableColumn { Id = 1, Name = "Item" });
var priceCol = new TableColumn { Id = 2, Name = "Price", TotalsRowFunction = TotalsRowFunctionValues.Sum };
tableCols.Append(priceCol);
var taxCol = new TableColumn { Id = 3, Name = "Tax" };
taxCol.Append(new CalculatedColumnFormula("[Price]*0.1"));
tableCols.Append(taxCol);
table.Append(tableCols);

table.Append(new TableStyleInfo
{
    Name = "TableStyleLight1", ShowFirstColumn = false, ShowLastColumn = false,
    ShowRowStripes = false, ShowColumnStripes = true
});

tablePart.Table = table;
tablePart.Table.Save();

doc.Dispose();

Console.Error.WriteLine("SCENARIO_PASS");
