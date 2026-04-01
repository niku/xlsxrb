// Generates an XLSX with a chart series that has line cap and join attributes.
var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData();
sheetData.Append(new Row(
    new Cell { CellReference = "A1", CellValue = new CellValue("10"), DataType = CellValues.Number },
    new Cell { CellReference = "A2", CellValue = new CellValue("20"), DataType = CellValues.Number }
));
worksheetPart.Worksheet = new Worksheet(sheetData);

var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
var chartPart = drawingsPart.AddNewPart<ChartPart>();

var chartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace();
var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();

var plotArea = new DocumentFormat.OpenXml.Drawing.Charts.PlotArea();
var lineChart = new DocumentFormat.OpenXml.Drawing.Charts.LineChart();
lineChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.Grouping { Val = DocumentFormat.OpenXml.Drawing.Charts.GroupingValues.Standard });
var ser = new DocumentFormat.OpenXml.Drawing.Charts.LineChartSeries();
ser.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 0u });
ser.Append(new DocumentFormat.OpenXml.Drawing.Charts.Order { Val = 0u });

// Add spPr with line cap and join
var spPr = new DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties();
var ln = new DocumentFormat.OpenXml.Drawing.Outline { Width = 38100, CapType = DocumentFormat.OpenXml.Drawing.LineCapValues.Square };
ln.Append(new DocumentFormat.OpenXml.Drawing.SolidFill(
    new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "0000FF" }));
ln.Append(new DocumentFormat.OpenXml.Drawing.Bevel());
spPr.Append(ln);
ser.Append(spPr);

var vals = new DocumentFormat.OpenXml.Drawing.Charts.Values();
var numRef = new DocumentFormat.OpenXml.Drawing.Charts.NumberReference();
numRef.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula("Sheet1!$A$1:$A$2"));
vals.Append(numRef);
ser.Append(vals);
lineChart.Append(ser);
lineChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 1u });
lineChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 2u });
plotArea.Append(lineChart);

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
        new DocumentFormat.OpenXml.Drawing.Charts.ChartReference { Id = drawingsPart.GetIdOfPart(chartPart) }
    ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" });
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
