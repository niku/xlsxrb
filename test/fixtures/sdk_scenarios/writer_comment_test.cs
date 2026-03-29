// Validates that an XLSX contains comments.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    var commentsPart = worksheetPart.WorksheetCommentsPart
        ?? throw new Exception("WorksheetCommentsPart is missing.");

    var comments = commentsPart.Comments ?? throw new Exception("Comments element is missing.");
    var commentList = comments.Elements<CommentList>().FirstOrDefault()
        ?? throw new Exception("CommentList is missing.");

    var allComments = commentList.Elements<Comment>().ToList();
    if (allComments.Count == 0)
        throw new Exception("No comments found.");

    var first = allComments.First();
    if (first.Reference?.Value != "A1")
        throw new Exception($"Expected first comment ref 'A1' but got '{first.Reference?.Value}'.");

    var text = first.Elements<CommentText>().FirstOrDefault()?.InnerText;
    if (text != "Hello comment")
        throw new Exception($"Expected comment text 'Hello comment' but got '{text}'.");
}
finally
{
    document.Dispose();
}
