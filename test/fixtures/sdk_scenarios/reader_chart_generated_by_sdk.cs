// Creates an XLSX with a bar chart.
var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Category")) },
            new Cell { CellReference = "B1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Value")) }),
    new Row(new Cell { CellReference = "A2", DataType = CellValues.InlineString,
        InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("A")) },
            new Cell { CellReference = "B2", CellValue = new CellValue("10") })
);
worksheetPart.Worksheet = new Worksheet(sheetData);

// Create chart.
var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
var chartPart = drawingsPart.AddNewPart<ChartPart>();

var chartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace();
var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
var title = new DocumentFormat.OpenXml.Drawing.Charts.Title();
var chartText = new DocumentFormat.OpenXml.Drawing.Charts.ChartText();
var richText = new DocumentFormat.OpenXml.Drawing.Charts.RichText();
richText.Append(new DocumentFormat.OpenXml.Drawing.BodyProperties());
richText.Append(new DocumentFormat.OpenXml.Drawing.ListStyle());
var para = new DocumentFormat.OpenXml.Drawing.Paragraph();
var run = new DocumentFormat.OpenXml.Drawing.Run();
run.Append(new DocumentFormat.OpenXml.Drawing.Text("Sales Data"));
para.Append(run);
richText.Append(para);
chartText.Append(richText);
title.Append(chartText);
chart.Append(title);

var plotArea = new DocumentFormat.OpenXml.Drawing.Charts.PlotArea();
plotArea.Append(new DocumentFormat.OpenXml.Drawing.Charts.Layout());
var barChart = new DocumentFormat.OpenXml.Drawing.Charts.BarChart();
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.BarDirection { Val = DocumentFormat.OpenXml.Drawing.Charts.BarDirectionValues.Column });
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.BarGrouping { Val = DocumentFormat.OpenXml.Drawing.Charts.BarGroupingValues.Clustered });

var series = new DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries();
series.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 0 });
series.Append(new DocumentFormat.OpenXml.Drawing.Charts.Order { Val = 0 });
barChart.Append(series);
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 1 });
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 2 });
plotArea.Append(barChart);

var catAx = new DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis();
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 1 });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Scaling(new DocumentFormat.OpenXml.Drawing.Charts.Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Delete { Val = false });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisPosition { Val = DocumentFormat.OpenXml.Drawing.Charts.AxisPositionValues.Bottom });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.CrossingAxis { Val = 2 });
plotArea.Append(catAx);

var valAx = new DocumentFormat.OpenXml.Drawing.Charts.ValueAxis();
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 2 });
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Scaling(new DocumentFormat.OpenXml.Drawing.Charts.Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Delete { Val = false });
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisPosition { Val = DocumentFormat.OpenXml.Drawing.Charts.AxisPositionValues.Left });
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.CrossingAxis { Val = 1 });
plotArea.Append(valAx);

chart.Append(plotArea);
chartSpace.Append(chart);
chartPart.ChartSpace = chartSpace;
chartPart.ChartSpace.Save();

// Build drawing with graphic frame pointing to chart.
var wsDr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
var anchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor();
var fromMarker = new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker();
fromMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("0"));
fromMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"));
fromMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("0"));
fromMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0"));
var toMarker = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker();
toMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("10"));
toMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"));
toMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("15"));
toMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0"));
anchor.Append(fromMarker);
anchor.Append(toMarker);

var gf = new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame();
gf.Macro = "";
var nvGfPr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties();
nvGfPr.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2, Name = "SalesChart" });
nvGfPr.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties());
gf.Append(nvGfPr);
gf.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.Transform());

var chartRelId = drawingsPart.GetIdOfPart(chartPart);
var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
var graphicData = new DocumentFormat.OpenXml.Drawing.GraphicData { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };
graphicData.InnerXml = $"<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"{chartRelId}\"/>";
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
