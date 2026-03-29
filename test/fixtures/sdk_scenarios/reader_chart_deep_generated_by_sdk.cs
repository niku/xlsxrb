// Creates an XLSX with a chart having multiple series, legend, data labels, and axis titles.
var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Cat")) },
            new Cell { CellReference = "B1", CellValue = new CellValue("10") },
            new Cell { CellReference = "C1", CellValue = new CellValue("20") }),
    new Row(new Cell { CellReference = "A2", DataType = CellValues.InlineString,
        InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Dog")) },
            new Cell { CellReference = "B2", CellValue = new CellValue("30") },
            new Cell { CellReference = "C2", CellValue = new CellValue("40") })
);
worksheetPart.Worksheet = new Worksheet(sheetData);

var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
var chartPart = drawingsPart.AddNewPart<ChartPart>();

var chartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace();
var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();

// Title
var title = new DocumentFormat.OpenXml.Drawing.Charts.Title();
var chartText = new DocumentFormat.OpenXml.Drawing.Charts.ChartText();
var richText = new DocumentFormat.OpenXml.Drawing.Charts.RichText();
richText.Append(new DocumentFormat.OpenXml.Drawing.BodyProperties());
richText.Append(new DocumentFormat.OpenXml.Drawing.ListStyle());
var titlePara = new DocumentFormat.OpenXml.Drawing.Paragraph();
var titleRun = new DocumentFormat.OpenXml.Drawing.Run();
titleRun.Append(new DocumentFormat.OpenXml.Drawing.Text("Multi Series"));
titlePara.Append(titleRun);
richText.Append(titlePara);
chartText.Append(richText);
title.Append(chartText);
chart.Append(title);

// Legend
var legend = new DocumentFormat.OpenXml.Drawing.Charts.Legend();
legend.Append(new DocumentFormat.OpenXml.Drawing.Charts.LegendPosition { Val = DocumentFormat.OpenXml.Drawing.Charts.LegendPositionValues.Bottom });
chart.Append(legend);

var plotArea = new DocumentFormat.OpenXml.Drawing.Charts.PlotArea();
plotArea.Append(new DocumentFormat.OpenXml.Drawing.Charts.Layout());

var barChart = new DocumentFormat.OpenXml.Drawing.Charts.BarChart();
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.BarDirection { Val = DocumentFormat.OpenXml.Drawing.Charts.BarDirectionValues.Column });
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.BarGrouping { Val = DocumentFormat.OpenXml.Drawing.Charts.BarGroupingValues.Clustered });

// Series 1
var ser1 = new DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries();
ser1.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 0 });
ser1.Append(new DocumentFormat.OpenXml.Drawing.Charts.Order { Val = 0 });
var dLbls1 = new DocumentFormat.OpenXml.Drawing.Charts.DataLabels();
dLbls1.Append(new DocumentFormat.OpenXml.Drawing.Charts.ShowValue { Val = true });
dLbls1.Append(new DocumentFormat.OpenXml.Drawing.Charts.ShowCategoryName { Val = false });
dLbls1.Append(new DocumentFormat.OpenXml.Drawing.Charts.ShowSeriesName { Val = false });
dLbls1.Append(new DocumentFormat.OpenXml.Drawing.Charts.ShowPercent { Val = false });
dLbls1.Append(new DocumentFormat.OpenXml.Drawing.Charts.ShowLegendKey { Val = false });
ser1.Append(dLbls1);
var catRef1 = new DocumentFormat.OpenXml.Drawing.Charts.CategoryAxisData();
var strRef1 = new DocumentFormat.OpenXml.Drawing.Charts.StringReference();
strRef1.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula("Sheet1!$A$1:$A$2"));
catRef1.Append(strRef1);
ser1.Append(catRef1);
var valRef1 = new DocumentFormat.OpenXml.Drawing.Charts.Values();
var numRef1 = new DocumentFormat.OpenXml.Drawing.Charts.NumberReference();
numRef1.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula("Sheet1!$B$1:$B$2"));
valRef1.Append(numRef1);
ser1.Append(valRef1);
barChart.Append(ser1);

// Series 2
var ser2 = new DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries();
ser2.Append(new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 1 });
ser2.Append(new DocumentFormat.OpenXml.Drawing.Charts.Order { Val = 1 });
var dLbls2 = new DocumentFormat.OpenXml.Drawing.Charts.DataLabels();
dLbls2.Append(new DocumentFormat.OpenXml.Drawing.Charts.ShowValue { Val = true });
dLbls2.Append(new DocumentFormat.OpenXml.Drawing.Charts.ShowCategoryName { Val = false });
dLbls2.Append(new DocumentFormat.OpenXml.Drawing.Charts.ShowSeriesName { Val = false });
dLbls2.Append(new DocumentFormat.OpenXml.Drawing.Charts.ShowPercent { Val = false });
dLbls2.Append(new DocumentFormat.OpenXml.Drawing.Charts.ShowLegendKey { Val = false });
ser2.Append(dLbls2);
var catRef2 = new DocumentFormat.OpenXml.Drawing.Charts.CategoryAxisData();
var strRef2 = new DocumentFormat.OpenXml.Drawing.Charts.StringReference();
strRef2.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula("Sheet1!$A$1:$A$2"));
catRef2.Append(strRef2);
ser2.Append(catRef2);
var valRef2 = new DocumentFormat.OpenXml.Drawing.Charts.Values();
var numRef2 = new DocumentFormat.OpenXml.Drawing.Charts.NumberReference();
numRef2.Append(new DocumentFormat.OpenXml.Drawing.Charts.Formula("Sheet1!$C$1:$C$2"));
valRef2.Append(numRef2);
ser2.Append(valRef2);
barChart.Append(ser2);

barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 1 });
barChart.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 2 });
plotArea.Append(barChart);

// Category Axis with title
var catAx = new DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis();
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 1 });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Scaling(new DocumentFormat.OpenXml.Drawing.Charts.Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Delete { Val = false });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisPosition { Val = DocumentFormat.OpenXml.Drawing.Charts.AxisPositionValues.Bottom });
catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.CrossingAxis { Val = 2 });
// CatAx title
var catAxTitle = new DocumentFormat.OpenXml.Drawing.Charts.Title();
var catAxTx = new DocumentFormat.OpenXml.Drawing.Charts.ChartText();
var catAxRt = new DocumentFormat.OpenXml.Drawing.Charts.RichText();
catAxRt.Append(new DocumentFormat.OpenXml.Drawing.BodyProperties());
catAxRt.Append(new DocumentFormat.OpenXml.Drawing.ListStyle());
var catAxPara = new DocumentFormat.OpenXml.Drawing.Paragraph();
var catAxRun = new DocumentFormat.OpenXml.Drawing.Run();
catAxRun.Append(new DocumentFormat.OpenXml.Drawing.Text("Category"));
catAxPara.Append(catAxRun);
catAxRt.Append(catAxPara);
catAxTx.Append(catAxRt);
catAxTitle.Append(catAxTx);
catAx.Append(catAxTitle);
plotArea.Append(catAx);

// Value Axis with title
var valAx = new DocumentFormat.OpenXml.Drawing.Charts.ValueAxis();
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisId { Val = 2 });
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Scaling(new DocumentFormat.OpenXml.Drawing.Charts.Orientation { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax }));
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.Delete { Val = false });
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.AxisPosition { Val = DocumentFormat.OpenXml.Drawing.Charts.AxisPositionValues.Left });
valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.CrossingAxis { Val = 1 });
// ValAx title
var valAxTitle = new DocumentFormat.OpenXml.Drawing.Charts.Title();
var valAxTx = new DocumentFormat.OpenXml.Drawing.Charts.ChartText();
var valAxRt = new DocumentFormat.OpenXml.Drawing.Charts.RichText();
valAxRt.Append(new DocumentFormat.OpenXml.Drawing.BodyProperties());
valAxRt.Append(new DocumentFormat.OpenXml.Drawing.ListStyle());
var valAxPara = new DocumentFormat.OpenXml.Drawing.Paragraph();
var valAxRun = new DocumentFormat.OpenXml.Drawing.Run();
valAxRun.Append(new DocumentFormat.OpenXml.Drawing.Text("Amount"));
valAxPara.Append(valAxRun);
valAxRt.Append(valAxPara);
valAxTx.Append(valAxRt);
valAxTitle.Append(valAxTx);
valAx.Append(valAxTitle);
plotArea.Append(valAx);

chart.Append(plotArea);
chartSpace.Append(chart);
chartPart.ChartSpace = chartSpace;
chartPart.ChartSpace.Save();

// Build drawing
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
nvGfPr.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2, Name = "MultiSeriesChart" });
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

Console.Error.WriteLine("SCENARIO_PASS");
