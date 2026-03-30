var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook(new Sheets(
        new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }
    ));

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
    worksheetPart.Worksheet = new Worksheet(new SheetData(
        new Row(new Cell { CellReference = "A1", DataType = CellValues.String, CellValue = new CellValue("data") })
        { RowIndex = 1 }
    ));

    var customPropsPart = document.AddCustomFilePropertiesPart();
    customPropsPart.Properties = new DocumentFormat.OpenXml.CustomProperties.Properties();

    var prop1 = new CustomDocumentProperty
    {
        FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
        PropertyId = 2,
        Name = "Project"
    };
    prop1.Append(new DocumentFormat.OpenXml.VariantTypes.VTLPWSTR("Alpha"));
    customPropsPart.Properties.Append(prop1);

    var prop2 = new CustomDocumentProperty
    {
        FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
        PropertyId = 3,
        Name = "Version"
    };
    prop2.Append(new DocumentFormat.OpenXml.VariantTypes.VTInt32("42"));
    customPropsPart.Properties.Append(prop2);

    var prop3 = new CustomDocumentProperty
    {
        FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}",
        PropertyId = 4,
        Name = "Active"
    };
    prop3.Append(new DocumentFormat.OpenXml.VariantTypes.VTBool("true"));
    customPropsPart.Properties.Append(prop3);

    document.Save();
}
finally
{
    document.Dispose();
}
