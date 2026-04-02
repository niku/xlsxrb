// Creates an XLSX with a pie chart having data labels with showLeaderLines and leaderLines spPr.
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
worksheetPart.Worksheet = new Worksheet(new SheetData());

var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
var chartPart = drawingsPart.AddNewPart<ChartPart>();

var chartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace();
chartSpace.InnerXml =
    "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" " +
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
    "<c:plotArea><c:pieChart>" +
    "<c:varyColors val=\"1\"/>" +
    "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>" +
    "<c:val><c:numRef><c:f>Sheet1!$A$1</c:f></c:numRef></c:val>" +
    "<c:dLbls>" +
    "<c:showVal val=\"1\"/>" +
    "<c:showLeaderLines val=\"1\"/>" +
    "<c:leaderLines><c:spPr>" +
    "<a:ln w=\"6350\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln>" +
    "</c:spPr></c:leaderLines>" +
    "</c:dLbls></c:ser>" +
    "</c:pieChart></c:plotArea></c:chart>";
chartPart.ChartSpace = chartSpace;
chartPart.ChartSpace.Save();

var wsDr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
var anchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor();
var fromM = new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker();
fromM.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("0"));
fromM.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"));
fromM.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("0"));
fromM.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0"));
var toM = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker();
toM.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("10"));
toM.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"));
toM.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("15"));
toM.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0"));
anchor.Append(fromM);
anchor.Append(toM);
var gf = new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame();
gf.Macro = "";
var nvGfPr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties();
nvGfPr.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2, Name = "Chart1" });
nvGfPr.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties());
gf.Append(nvGfPr);
gf.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.Transform());
var chartRelId = drawingsPart.GetIdOfPart(chartPart);
var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData
    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };
graphicData.InnerXml = $"<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{chartRelId}\"/>";
var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
graphic.Append(graphicData);
gf.Append(graphic);
anchor.Append(gf);
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());
wsDr.Append(anchor);
drawingsPart.WorksheetDrawing = wsDr;
drawingsPart.WorksheetDrawing.Save();

var drawingRelId = worksheetPart.GetIdOfPart(drawingsPart);
worksheetPart.Worksheet.Append(new Drawing { Id = drawingRelId });
worksheetPart.Worksheet.Save();

var sheetsEl = workbookPart.Workbook.AppendChild(new Sheets());
sheetsEl.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();

Console.Error.WriteLine("SCENARIO_PASS");
