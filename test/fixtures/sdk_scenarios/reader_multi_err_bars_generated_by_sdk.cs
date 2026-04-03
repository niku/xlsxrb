// Generates an XLSX with a scatter chart having two errBars (x and y directions) on a single series.
var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
try
{
    var wbPart = document.AddWorkbookPart();
    wbPart.Workbook = new Workbook(new Sheets(
        new Sheet { Name = "Sheet1", SheetId = 1, Id = "rId1" }
    ));
    var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");
    wsPart.Worksheet = new Worksheet(new SheetData(
        new Row(new Cell { CellReference = "A1", DataType = CellValues.Number, CellValue = new CellValue("1") },
                new Cell { CellReference = "B1", DataType = CellValues.Number, CellValue = new CellValue("2") })
    ));

    var drawingsPart = wsPart.AddNewPart<DrawingsPart>();
    wsPart.Worksheet.Append(new Drawing { Id = wsPart.GetIdOfPart(drawingsPart) });
    drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
    drawingsPart.WorksheetDrawing.InnerXml = @"
<xdr:twoCellAnchor xmlns:xdr=""http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing""
                   xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""
                   xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
  <xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
  <xdr:graphicFrame macro="""">
    <xdr:nvGraphicFramePr>
      <xdr:cNvPr id=""2"" name=""Chart 1""/>
      <xdr:cNvGraphicFramePr/>
    </xdr:nvGraphicFramePr>
    <xdr:xfrm><a:off x=""0"" y=""0""/><a:ext cx=""0"" cy=""0""/></xdr:xfrm>
    <a:graphic>
      <a:graphicData uri=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
        <c:chart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"" r:id=""rId1""/>
      </a:graphicData>
    </a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData/>
</xdr:twoCellAnchor>";

    var chartPart = drawingsPart.AddNewPart<ChartPart>("rId1");
    chartPart.ChartSpace = new DocumentFormat.OpenXml.Drawing.Charts.ChartSpace();
    chartPart.ChartSpace.InnerXml = @"
<c:chart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart"">
  <c:plotArea>
    <c:layout/>
    <c:scatterChart>
      <c:scatterStyle val=""lineMarker""/>
      <c:varyColors val=""0""/>
      <c:ser>
        <c:idx val=""0""/>
        <c:order val=""0""/>
        <c:xVal><c:numRef><c:f>Sheet1!$A$1</c:f></c:numRef></c:xVal>
        <c:yVal><c:numRef><c:f>Sheet1!$B$1</c:f></c:numRef></c:yVal>
        <c:errBars>
          <c:errDir val=""x""/>
          <c:errBarType val=""both""/>
          <c:errValType val=""fixedVal""/>
          <c:val val=""0.5""/>
        </c:errBars>
        <c:errBars>
          <c:errDir val=""y""/>
          <c:errBarType val=""both""/>
          <c:errValType val=""fixedVal""/>
          <c:val val=""1.0""/>
        </c:errBars>
      </c:ser>
      <c:axId val=""1""/>
      <c:axId val=""2""/>
    </c:scatterChart>
    <c:valAx><c:axId val=""1""/><c:scaling><c:orientation val=""minMax""/></c:scaling><c:delete val=""0""/><c:axPos val=""b""/><c:crossAx val=""2""/></c:valAx>
    <c:valAx><c:axId val=""2""/><c:scaling><c:orientation val=""minMax""/></c:scaling><c:delete val=""0""/><c:axPos val=""l""/><c:crossAx val=""1""/></c:valAx>
  </c:plotArea>
</c:chart>";

    document.Save();
    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
