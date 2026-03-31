// Creates an XLSX with a chart having legend entries (one deleted).
var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", CellValue = new CellValue("10") }),
    new Row(new Cell { CellReference = "A2", CellValue = new CellValue("20") })
);
worksheetPart.Worksheet = new Worksheet(sheetData);

var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
var chartPart = drawingsPart.AddNewPart<ChartPart>();

var chartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace();
var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();

// Legend with entries
var legend = new DocumentFormat.OpenXml.Drawing.Charts.Legend();
legend.Append(new DocumentFormat.OpenXml.Drawing.Charts.LegendPosition { Val = DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues.Right });
var legendEntry = new DocumentFormat.OpenXml.Drawing.Charts.LegendEntry();
legendEntry.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 0u });
legendEntry.Append(new DocumentFormat.OpenXml.Drawing.Charts.Delete { Val = true });
legend.Append(legendEntry);
chart.Append(legend);

var plotArea = new DocumentFormat.OpenXml.Drawing.Charts.PlotArea();
plotArea.Append(new DocumentFormat.OpenXml.Drawing.Charts.Layout());

var barChart = new DocumentFormat.OpenXml.Drawing.Charts.BarChart();
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.BarDirection { Val = DocumentFormat.OpenXml.Drawing.Charts.BarDirectionValues.Column });
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.BarGrouping { Val = DocumentFormat.OpenXml.Drawing.Charts.BarGroupingValues.Clustered });

var series1 = new DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries();
series1.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 0u });
series1.Append(new DocumentFormat.OpenXml.Drawing.Charts.Order { Val = 0u });
barChart.Append(series1);

var series2 = new DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries();
series2.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 1u });
series2.Append(new DocumentFormat.OpenXml.Drawing.Charts.Order { Val = 1u });
barChart.Append(series2);

plotArea.Append(barChart);

var catAx = new DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis();
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 1u });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Scaling(
    new DocumentFormat.OpenXml.Drawing.Charts.Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Delete { Val = false });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisPosition { Val = DocumentFormat.OpenXml.Drawing.Charts.AxisPositionValues.Bottom });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.CrossingAxis { Val = 2u });
plotArea.Append(catAx);

var valAx = new DocumentFormat.OpenXml.Drawing.Charts.ValueAxis();
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 2u });
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Scaling(
    new DocumentFormat.OpenXml.Drawing.Charts.Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Delete { Val = false });
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisPosition { Val = DocumentFormat.OpenXml.Drawing.Charts.AxisPositionValues.Left });
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.CrossingAxis { Val = 1u });
plotArea.Append(valAx);

chart.Append(plotArea);
chartSpace.Append(chart);
chartPart.ChartSpace = chartSpace;

// Drawing anchor
var wsDr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
var anchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor();
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("0"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("0"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")));
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("10"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("15"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")));
var gf = new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame { Macro = "" };
gf.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2u, Name = "Chart 1" },
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));
var transform = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Transform();
transform.Append(new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 });
transform.Append(new DocumentFormat.OpenXml.Drawing.Extents { Cx = 0, Cy = 0 });
gf.Append(transform);
var graphic = new DocumentFormat.OpenXml.Drawing.Graphic();
var gd = new DocumentFormat.OpenXml.Drawing.GraphicData { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" };
gd.Append(new DocumentFormat.OpenXml.Drawing.Charts.ChartReference { Id = drawingsPart.GetIdOfPart(chartPart) });
graphic.Append(gd);
gf.Append(graphic);
anchor.Append(gf);
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());
wsDr.Append(anchor);
drawingsPart.WorksheetDrawing = wsDr;

worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });

var sheets = new Sheets();
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1u, Name = "Sheet1" });
workbookPart.Workbook.Append(sheets);

doc.Dispose();
