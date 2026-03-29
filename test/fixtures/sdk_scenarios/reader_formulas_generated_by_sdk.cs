// Creates an XLSX with shared and array formulas.
var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData();

// Row 1: A1=10, B1=shared formula master (A1*2)
var row1 = new Row { RowIndex = 1 };
row1.Append(new Cell { CellReference = "A1", CellValue = new CellValue("10") });
var b1 = new Cell { CellReference = "B1", CellValue = new CellValue("20") };
var f1 = new CellFormula("A1*2") { FormulaType = CellFormulaValues.Shared, Reference = "B1:B2", SharedIndex = 0 };
b1.CellFormula = f1;
row1.Append(b1);
sheetData.Append(row1);

// Row 2: A2=20, B2=shared formula secondary
var row2 = new Row { RowIndex = 2 };
row2.Append(new Cell { CellReference = "A2", CellValue = new CellValue("20") });
var b2 = new Cell { CellReference = "B2", CellValue = new CellValue("40") };
var f2 = new CellFormula { FormulaType = CellFormulaValues.Shared, SharedIndex = 0 };
b2.CellFormula = f2;
row2.Append(b2);
sheetData.Append(row2);

// Row 3: C3=array formula
var row3 = new Row { RowIndex = 3 };
var c3 = new Cell { CellReference = "C3", CellValue = new CellValue("70") };
var f3 = new CellFormula("SUM(A1:A2*B1:B2)") { FormulaType = CellFormulaValues.Array, Reference = "C3" };
c3.CellFormula = f3;
row3.Append(c3);
sheetData.Append(row3);

worksheetPart.Worksheet = new Worksheet(sheetData);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
