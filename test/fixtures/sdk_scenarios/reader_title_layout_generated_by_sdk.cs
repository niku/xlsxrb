// Generates an XLSX with chart title and axis title layout positioning for reader interoperability testing.
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
  <xdr:to><xdr:col>10</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>15</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
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
  <c:title>
    <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>SDK Title</a:t></a:r></a:p></c:rich></c:tx>
    <c:layout>
      <c:manualLayout>
        <c:x val=""0.15""/>
        <c:y val=""0.03""/>
        <c:w val=""0.7""/>
        <c:h val=""0.06""/>
      </c:manualLayout>
    </c:layout>
    <c:overlay val=""0""/>
  </c:title>
  <c:plotArea>
    <c:layout/>
    <c:barChart>
      <c:barDir val=""col""/>
      <c:grouping val=""clustered""/>
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
        <c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>X Axis</a:t></a:r></a:p></c:rich></c:tx>
        <c:layout>
          <c:manualLayout>
            <c:x val=""0.25""/>
            <c:y val=""0.88""/>
          </c:manualLayout>
        </c:layout>
        <c:overlay val=""0""/>
      </c:title>
      <c:crossAx val=""2""/>
    </c:catAx>
    <c:valAx>
      <c:axId val=""2""/>
      <c:scaling><c:orientation val=""minMax""/></c:scaling>
      <c:delete val=""0""/>
      <c:axPos val=""l""/>
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
