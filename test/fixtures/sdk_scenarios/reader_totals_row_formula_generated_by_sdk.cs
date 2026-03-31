// Creates an XLSX with a table column that has a totalsRowFormula element.
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
row1.Append(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
    InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Item")) });
row1.Append(new Cell { CellReference = "B1", DataType = CellValues.InlineString,
    InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Price")) });
sd.Append(row1);
ws.Append(sd);

var tableParts = new TableParts { Count = 1 };
tableParts.Append(new TablePart { Id = "rId2" });
ws.Append(tableParts);
wsPart.Worksheet = ws;
wsPart.Worksheet.Save();

var tablePart = wsPart.AddNewPart<TableDefinitionPart>("rId2");
var table = new Table
{
    Id = 1, Name = "PriceTable", DisplayName = "PriceTable",
    Reference = "A1:B3", TotalsRowCount = 1
};
table.Append(new AutoFilter { Reference = "A1:B3" });

var tableCols = new TableColumns { Count = 2 };
tableCols.Append(new TableColumn { Id = 1, Name = "Item" });
var priceCol = new TableColumn
{
    Id = 2, Name = "Price", TotalsRowFunction = TotalsRowFunctionValues.Custom
};
priceCol.Append(new TotalsRowFormula("SUBTOTAL(109,[Price])"));
tableCols.Append(priceCol);
table.Append(tableCols);
table.Append(new TableStyleInfo { Name = "TableStyleMedium2", ShowFirstColumn = false,
    ShowLastColumn = false, ShowRowStripes = true, ShowColumnStripes = false });
tablePart.Table = table;
tablePart.Table.Save();

doc.Dispose();
Console.Error.WriteLine("SCENARIO_PASS");
