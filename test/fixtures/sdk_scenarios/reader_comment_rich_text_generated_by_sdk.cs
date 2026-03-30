var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    var sheetData = new SheetData(
        new Row(
            new Cell
            {
                CellReference = "A1",
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text("data"))
            }
        )
    );
    worksheetPart.Worksheet = new Worksheet(sheetData);

    // Add comments part
    var commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();
    var authors = new Authors(new Author("TestAuthor"));

    // Build rich text comment: bold "Important" + plain " note text"
    var commentText = new CommentText(
        new Run(
            new RunProperties(
                new Bold(),
                new FontSize() { Val = 11.0 },
                new RunFont() { Val = "Calibri" }
            ),
            new Text("Important")
        ),
        new Run(
            new Text(" note text")
        )
    );

    var commentList = new CommentList(
        new Comment(commentText)
        {
            Reference = "A1",
            AuthorId = 0
        }
    );

    commentsPart.Comments = new Comments(authors, commentList);
    commentsPart.Comments.Save();

    // Add VML drawing part for comment shapes
    var vmlPart = worksheetPart.AddNewPart<VmlDrawingPart>();
    using (var writer = new System.IO.StreamWriter(vmlPart.GetStream()))
    {
        writer.Write(@"<xml xmlns:v=""urn:schemas-microsoft-com:vml"" xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">
<o:shapelayout v:ext=""edit""><o:idmap v:ext=""edit"" data=""1""/></o:shapelayout>
<v:shapetype id=""_x0000_t202"" coordsize=""21600,21600"" o:spt=""202"" path=""m,l,21600r21600,l21600,xe"">
<v:stroke joinstyle=""miter""/><v:path gradientshapeok=""t"" o:connecttype=""rect""/>
</v:shapetype>
<v:shape id=""_x0000_s1025"" type=""#_x0000_t202"" style=""position:absolute;margin-left:80pt;margin-top:2pt;width:120pt;height:60pt;z-index:1"" fillcolor=""#ffffe1"" o:insetmode=""auto"">
<v:fill color2=""#ffffe1""/><v:shadow on=""t"" color=""black"" obscured=""t""/>
<v:textbox style=""mso-direction-alt:auto""><div style=""text-align:left""></div></v:textbox>
<x:ClientData ObjectType=""Note""><x:Anchor>1, 15, 0, 10, 3, 15, 4, 4</x:Anchor><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData>
</v:shape></xml>");
    }

    // Add legacy drawing reference
    worksheetPart.Worksheet.AppendChild(
        new LegacyDrawing() { Id = worksheetPart.GetIdOfPart(vmlPart) }
    );

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
