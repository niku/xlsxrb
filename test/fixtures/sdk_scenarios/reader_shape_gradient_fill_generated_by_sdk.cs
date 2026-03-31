// Generates an XLSX with a shape that has a gradient fill (a:gradFill with gsLst and lin).
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

var spPr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties();
spPr.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry(
    new DocumentFormat.OpenXml.Drawing.AdjustValueList()) { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle });

var gsLst = new DocumentFormat.OpenXml.Drawing.GradientStopList();
var gs1 = new DocumentFormat.OpenXml.Drawing.GradientStop { Position = 0 };
gs1.Append(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "00FF00" });
gsLst.Append(gs1);
var gs2 = new DocumentFormat.OpenXml.Drawing.GradientStop { Position = 50000 };
gs2.Append(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FFFF00" });
gsLst.Append(gs2);
var gs3 = new DocumentFormat.OpenXml.Drawing.GradientStop { Position = 100000 };
gs3.Append(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FF00FF" });
gsLst.Append(gs3);

var gradFill = new DocumentFormat.OpenXml.Drawing.GradientFill();
gradFill.Append(gsLst);
gradFill.Append(new DocumentFormat.OpenXml.Drawing.LinearGradientFill { Angle = 2700000, Scaled = false });
spPr.Append(gradFill);
sp.Append(spPr);

var txBody = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody();
txBody.Append(new DocumentFormat.OpenXml.Drawing.BodyProperties());
txBody.Append(new DocumentFormat.OpenXml.Drawing.ListStyle());
var run = new DocumentFormat.OpenXml.Drawing.Run();
run.Append(new DocumentFormat.OpenXml.Drawing.RunProperties { Language = "en-US" });
run.Append(new DocumentFormat.OpenXml.Drawing.Text("Gradient Shape"));
txBody.Append(new DocumentFormat.OpenXml.Drawing.Paragraph(run));
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
