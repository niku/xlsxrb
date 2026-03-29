var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var columns = new Columns(
    new Column { Min = 1, Max = 1, Width = 10, CustomWidth = true },
    new Column { Min = 2, Max = 2, Width = 8, Hidden = true, CustomWidth = true },
    new Column { Min = 3, Max = 3, Width = 12, OutlineLevel = 2, Collapsed = true, CustomWidth = true }
);
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("hello")) }) { RowIndex = 1 },
    new Row(new Cell { CellReference = "A2", DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("outline")) }) { RowIndex = 2, OutlineLevel = 1 },
    new Row() { RowIndex = 3, Collapsed = true }
);
worksheetPart.Worksheet = new Worksheet(columns, sheetData);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
