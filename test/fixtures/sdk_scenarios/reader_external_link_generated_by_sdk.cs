// Creates an XLSX with an external link referencing another workbook.
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var wbPart = doc.AddWorkbookPart();
wbPart.Workbook = new Workbook();

var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");
wsPart.Worksheet = new Worksheet(new SheetData());
wsPart.Worksheet.Save();

// Create external workbook part
var extPart = wbPart.AddNewPart<ExternalWorkbookPart>();
var extLink = new ExternalLink();
var extBook = new ExternalBook();
extBook.Id = "rId1";
var sheetNames = new SheetNames();
sheetNames.Append(new SheetName { Val = "RemoteSheet1" });
sheetNames.Append(new SheetName { Val = "RemoteSheet2" });
extBook.Append(sheetNames);
extLink.Append(extBook);
extPart.ExternalLink = extLink;
extPart.ExternalLink.Save();

// Add external rels pointing to target
extPart.AddExternalRelationship(
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath",
    new Uri("ExtBook.xlsx", UriKind.Relative),
    "rId1"
);

// Workbook external references
var extRid = wbPart.GetIdOfPart(extPart);
wbPart.Workbook.Append(new ExternalReferences(
    new ExternalReference { Id = extRid }
));

var sheetsEl = wbPart.Workbook.AppendChild(new Sheets());
sheetsEl.Append(new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" });
wbPart.Workbook.Save();
doc.Dispose();

Console.Error.WriteLine("SCENARIO_PASS");
