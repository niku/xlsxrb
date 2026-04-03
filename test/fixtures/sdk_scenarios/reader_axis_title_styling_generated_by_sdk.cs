// Generates an XLSX with axis titles having font formatting and spPr styling.
var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
try
{
    var wbPart = document.AddWorkbookPart();
    wbPart.Workbook = new Workbook(new Sheets(
        new Sheet { Name = "Sheet1", SheetId = 1, Id = "rId1" }
    ));
    var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");
    wsPart.Worksheet = new Worksheet(new SheetData(
        new Row(new Cell { CellReference = "A1", DataType = CellValues.Number, CellValue = new CellValue("10") })
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
<c:chart xmlns:c=""http://schemas.openxmlformats.org/drawingml/2006/chart""
         xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"">
  <c:plotArea>
    <c:layout/>
    <c:barChart>
      <c:barDir val=""col""/>
      <c:grouping val=""clustered""/>
      <c:varyColors val=""0""/>
      <c:ser>
        <c:idx val=""0""/>
        <c:order val=""0""/>
        <c:val><c:numRef><c:f>Sheet1!$A$1</c:f></c:numRef></c:val>
      </c:ser>
      <c:axId val=""1""/>
      <c:axId val=""2""/>
    </c:barChart>
    <c:catAx>
      <c:axId val=""1""/>
      <c:scaling><c:orientation val=""minMax""/></c:scaling>
      <c:delete val=""0""/>
      <c:axPos val=""b""/>
      <c:title>
        <c:tx><c:rich><a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:rPr b=""1"" sz=""1200""><a:solidFill><a:srgbClr val=""0000FF""/></a:solidFill></a:rPr><a:t>Category</a:t></a:r></a:p>
        </c:rich></c:tx>
        <c:overlay val=""0""/>
        <c:spPr><a:solidFill><a:srgbClr val=""EEEEFF""/></a:solidFill><a:ln w=""6350""><a:solidFill><a:srgbClr val=""0000CC""/></a:solidFill><a:prstDash val=""dot""/></a:ln></c:spPr>
      </c:title>
      <c:crossAx val=""2""/>
    </c:catAx>
    <c:valAx>
      <c:axId val=""2""/>
      <c:scaling><c:orientation val=""minMax""/></c:scaling>
      <c:delete val=""0""/>
      <c:axPos val=""l""/>
      <c:title>
        <c:tx><c:rich><a:bodyPr/><a:lstStyle/>
          <a:p><a:r><a:rPr i=""1"" sz=""1000""/><a:t>Value</a:t></a:r></a:p>
        </c:rich></c:tx>
        <c:overlay val=""0""/>
        <c:spPr><a:noFill/></c:spPr>
      </c:title>
      <c:crossAx val=""1""/>
    </c:valAx>
  </c:plotArea>
</c:chart>";

    document.Save();
    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
