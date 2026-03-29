// Creates an XLSX with shapes having preset geometry and text body.
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var wbPart = doc.AddWorkbookPart();
wbPart.Workbook = new Workbook();

var wsPart = wbPart.AddNewPart<WorksheetPart>();
wsPart.Worksheet = new Worksheet(new SheetData());

var drawingsPart = wsPart.AddNewPart<DrawingsPart>();
var wsDr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();

// Shape 1: ellipse with text
var anchor1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor();
var from1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker();
from1.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("1"));
from1.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"));
from1.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("2"));
from1.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0"));
var to1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker();
to1.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("4"));
to1.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"));
to1.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("6"));
to1.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0"));
anchor1.Append(from1);
anchor1.Append(to1);

var sp1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape();
var nvSpPr1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualShapeProperties();
nvSpPr1.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2, Name = "Oval 1" });
nvSpPr1.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualShapeDrawingProperties());
sp1.Append(nvSpPr1);

var spPr1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties();
spPr1.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse });
sp1.Append(spPr1);

var txBody1 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody();
txBody1.Append(new DocumentFormat.OpenXml.Drawing.BodyProperties());
txBody1.Append(new DocumentFormat.OpenXml.Drawing.ListStyle());
var p1 = new DocumentFormat.OpenXml.Drawing.Paragraph();
var r1 = new DocumentFormat.OpenXml.Drawing.Run();
r1.Append(new DocumentFormat.OpenXml.Drawing.Text("ShapeText"));
p1.Append(r1);
txBody1.Append(p1);
sp1.Append(txBody1);

anchor1.Append(sp1);
anchor1.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());
wsDr.Append(anchor1);

// Shape 2: roundRect no text
var anchor2 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor();
var from2 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker();
from2.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("5"));
from2.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"));
from2.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("0"));
from2.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0"));
var to2 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker();
to2.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("8"));
to2.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"));
to2.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("3"));
to2.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0"));
anchor2.Append(from2);
anchor2.Append(to2);

var sp2 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape();
var nvSpPr2 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualShapeProperties();
nvSpPr2.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 3, Name = "RoundRect 1" });
nvSpPr2.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualShapeDrawingProperties());
sp2.Append(nvSpPr2);
var spPr2 = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties();
spPr2.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RoundRectangle });
sp2.Append(spPr2);
sp2.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody(
    new DocumentFormat.OpenXml.Drawing.BodyProperties(),
    new DocumentFormat.OpenXml.Drawing.ListStyle(),
    new DocumentFormat.OpenXml.Drawing.Paragraph()));

anchor2.Append(sp2);
anchor2.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());
wsDr.Append(anchor2);

drawingsPart.WorksheetDrawing = wsDr;
drawingsPart.WorksheetDrawing.Save();

var drawingRelId = wsPart.GetIdOfPart(drawingsPart);
wsPart.Worksheet.Append(new Drawing { Id = drawingRelId });
wsPart.Worksheet.Save();

var sheetsEl = wbPart.Workbook.AppendChild(new Sheets());
sheetsEl.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1, Name = "Sheet1" });
wbPart.Workbook.Save();
doc.Dispose();

Console.Error.WriteLine("SCENARIO_PASS");
