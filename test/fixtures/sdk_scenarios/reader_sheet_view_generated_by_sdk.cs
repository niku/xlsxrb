var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetViews = new SheetViews(
    new SheetView(
        new Pane { VerticalSplit = 2, HorizontalSplit = 1, TopLeftCell = "B3", ActivePane = PaneValues.BottomRight, State = PaneStateValues.Frozen },
        new Selection { ActiveCell = "D4", SequenceOfReferences = new ListValue<StringValue>(new[] { StringValue.FromString("D4") }) }
    ) { ShowGridLines = false, ZoomScale = 120, WorkbookViewId = 0 }
);
worksheetPart.Worksheet = new Worksheet(sheetViews, new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("hello")) })
));
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
