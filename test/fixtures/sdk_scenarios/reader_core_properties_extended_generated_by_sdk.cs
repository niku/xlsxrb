var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    // Add core properties
    document.AddCoreFilePropertiesPart();
    var cpNs = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    var dcNs = "http://purl.org/dc/elements/1.1/";
    var dctermsNs = "http://purl.org/dc/terms/";
    var xsiNs = "http://www.w3.org/2001/XMLSchema-instance";
    var xdoc = new System.Xml.Linq.XDocument(
        new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("coreProperties", cpNs),
            new System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xmlns + "dc", dcNs),
            new System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xmlns + "dcterms", dctermsNs),
            new System.Xml.Linq.XAttribute(System.Xml.Linq.XNamespace.Xmlns + "xsi", xsiNs),
            new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("title", dcNs), "SDK Title"),
            new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("subject", dcNs), "SDK Subject"),
            new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("creator", dcNs), "SDK Creator"),
            new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("keywords", cpNs), "sdk, test"),
            new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("description", dcNs), "SDK Description"),
            new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("lastModifiedBy", cpNs), "SDK Editor"),
            new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("revision", cpNs), "5"),
            new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("category", cpNs), "Testing"),
            new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("contentStatus", cpNs), "Final"),
            new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("language", dcNs), "ja-JP")
        )
    );
    xdoc.Save(document.CoreFilePropertiesPart.GetStream(System.IO.FileMode.Create));

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    var sheetData = new SheetData(
        new Row(new Cell { CellReference = "A1", CellValue = new CellValue("test"), DataType = CellValues.InlineString, InlineString = new InlineString(new Text("test")) })
    );
    worksheetPart.Worksheet = new Worksheet(sheetData);

    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
    sheets.Append(new Sheet
    {
        Id = workbookPart.GetIdOfPart(worksheetPart),
        SheetId = 1,
        Name = "Sheet1"
    });

    workbookPart.Workbook.Save();
    worksheetPart.Worksheet.Save();
}
finally
{
    document.Dispose();
}
