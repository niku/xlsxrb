// Creates an XLSX with a shape having solidFill color.
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

var anchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor();
var from = new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker();
from.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("0"));
from.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"));
from.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("0"));
from.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0"));
var to = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker();
to.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("5"));
to.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"));
to.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("5"));
to.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0"));
anchor.Append(from);
anchor.Append(to);

var sp = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape();
var nvSpPr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualShapeProperties();
nvSpPr.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2, Name = "Rect 1" });
nvSpPr.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualShapeDrawingProperties());
sp.Append(nvSpPr);

var spPr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties();
spPr.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle });
var solidFill = new DocumentFormat.OpenXml.Drawing.SolidFill();
solidFill.Append(new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "FF0000" });
spPr.Append(solidFill);
sp.Append(spPr);

sp.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody(
    new DocumentFormat.OpenXml.Drawing.BodyProperties(),
    new DocumentFormat.OpenXml.Drawing.ListStyle(),
    new DocumentFormat.OpenXml.Drawing.Paragraph()));

anchor.Append(sp);
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());
wsDr.Append(anchor);
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
