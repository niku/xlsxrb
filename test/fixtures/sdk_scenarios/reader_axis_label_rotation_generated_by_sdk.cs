// Generates an XLSX with a chart where catAx has txPr with rot attribute on bodyPr.
var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
worksheetPart.Worksheet = new Worksheet(new SheetData(
    new Row(new Cell { CellReference = "A1", CellValue = new CellValue("Cat1"), DataType = CellValues.String }),
    new Row(new Cell { CellReference = "B1", CellValue = new CellValue("10") })
));

var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
var chartPart = drawingsPart.AddNewPart<ChartPart>();

var chartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace();
chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
chartSpace.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();
var plotArea = new DocumentFormat.OpenXml.Drawing.Charts.PlotArea();
plotArea.Append(new DocumentFormat.OpenXml.Drawing.Charts.Layout());

var barChart = new DocumentFormat.OpenXml.Drawing.Charts.BarChart();
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.BarDirection { Val = DocumentFormat.OpenXml.Drawing.Charts.BarDirectionValues.Column });
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.BarGrouping { Val = DocumentFormat.OpenXml.Drawing.Charts.BarGroupingValues.Clustered });
var ser = new DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries();
ser.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 0u });
ser.Append(new DocumentFormat.OpenXml.Drawing.Charts.Order { Val = 0u });
barChart.Append(ser);
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 1u });
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 2u });
plotArea.Append(barChart);

// Category axis with txPr rotation
var catAx = new DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis();
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 1u });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Scaling(
    new DocumentFormat.OpenXml.Drawing.Charts.Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Delete { Val = false });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisPosition { Val = DocumentFormat.OpenXml.Drawing.Charts.AxisPositionValues.Bottom });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.TickLabelPosition { Val = DocumentFormat.OpenXml.Drawing.Charts.TickLabelPositionValues.NextTo });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.TextProperties(
    new DocumentFormat.OpenXml.Drawing.BodyProperties { Rotation = -2700000 },
    new DocumentFormat.OpenXml.Drawing.ListStyle(),
    new DocumentFormat.OpenXml.Drawing.Paragraph(
        new DocumentFormat.OpenXml.Drawing.ParagraphProperties(
            new DocumentFormat.OpenXml.Drawing.DefaultRunProperties()),
        new DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties())));
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.CrossingAxis { Val = 2u });
plotArea.Append(catAx);

// Value axis
var valAx = new DocumentFormat.OpenXml.Drawing.Charts.ValueAxis();
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 2u });
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Scaling(
    new DocumentFormat.OpenXml.Drawing.Charts.Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Delete { Val = false });
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisPosition { Val = DocumentFormat.OpenXml.Drawing.Charts.AxisPositionValues.Left });
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.CrossingAxis { Val = 1u });
plotArea.Append(valAx);

chart.Append(plotArea);
chart.Append(new DocumentFormat.OpenXml.Drawing.Charts.Legend(
    new DocumentFormat.OpenXml.Drawing.Charts.LegendPosition { Val = DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues.Right }));
chartSpace.Append(chart);
chartPart.ChartSpace = chartSpace;

// Wire up drawing
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

var graphicFrame = new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame();
graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2u, Name = "Chart 1" },
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));
graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.Transform(
    new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
    new DocumentFormat.OpenXml.Drawing.Extents { Cx = 0, Cy = 0 }));
var graphic = new DocumentFormat.OpenXml.Drawing.Graphic(
    new DocumentFormat.OpenXml.Drawing.GraphicData(
        new DocumentFormat.OpenXml.Drawing.Charts.ChartReference { Id = drawingsPart.GetIdOfPart(chartPart) })
    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" });
graphicFrame.Append(graphic);
anchor.Append(graphicFrame);
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());
wsDr.Append(anchor);
drawingsPart.WorksheetDrawing = wsDr;

worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });

var sheets = new Sheets();
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1u, Name = "Sheet1" });
workbookPart.Workbook.Append(sheets);

doc.Dispose();
