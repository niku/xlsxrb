// Generates an XLSX with a shape that has text paragraph alignment.
var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
worksheetPart.Worksheet = new Worksheet(new SheetData(
    new Row(new Cell { CellReference = "A1", CellValue = new CellValue("test") })
));

var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

var wsDr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
var anchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor();
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("0"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("0"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")));
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("5"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("5"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")));

var sp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape();

sp.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualShapeProperties(
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2u, Name = "Shape 1" },
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualShapeDrawingProperties()));

sp.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties(
    new DocumentFormat.OpenXml.Drawing.PresetGeometry(
        new DocumentFormat.OpenXml.Drawing.AdjustValueList()) { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }));

// TextBody with paragraph alignment
var txBody = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody();
txBody.Append(new DocumentFormat.OpenXml.Drawing.BodyProperties());
txBody.Append(new DocumentFormat.OpenXml.Drawing.ListStyle());

var pPr = new DocumentFormat.OpenXml.Drawing.ParagraphProperties { Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center };

var run = new DocumentFormat.OpenXml.Drawing.Run();
run.Append(new DocumentFormat.OpenXml.Drawing.Text("Centered"));
var para = new DocumentFormat.OpenXml.Drawing.Paragraph();
para.Append(pPr);
para.Append(run);
txBody.Append(para);
sp.Append(txBody);

sp.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeStyle(
    new DocumentFormat.OpenXml.Drawing.LineReference(
        new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent1 }) { Index = 2u },
    new DocumentFormat.OpenXml.Drawing.FillReference(
        new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent1 }) { Index = 1u },
    new DocumentFormat.OpenXml.Drawing.EffectReference(
        new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.Accent1 }) { Index = 0u },
    new DocumentFormat.OpenXml.Drawing.FontReference(
        new DocumentFormat.OpenXml.Drawing.SchemeColor { Val = DocumentFormat.OpenXml.Drawing.SchemeColorValues.Dark1 }) { Index = DocumentFormat.OpenXml.Drawing.FontCollectionIndexValues.Minor }
));

anchor.Append(sp);
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());
wsDr.Append(anchor);
drawingsPart.WorksheetDrawing = wsDr;

worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });

var sheets = new Sheets();
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1u, Name = "Sheet1" });
workbookPart.Workbook.Append(sheets);

doc.Dispose();
