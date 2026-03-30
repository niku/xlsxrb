var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var validationErrors = validator.Validate(document).Take(10).ToList();
    if (validationErrors.Any())
    {
        var message = string.Join(Environment.NewLine, validationErrors.Select(e => e.Description));
        throw new Exception($"OpenXmlValidator reported errors:{Environment.NewLine}{message}");
    }

    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet definition is missing.");

    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id!.Value!);

    // Find comments part
    var commentsPart = worksheetPart.WorksheetCommentsPart
        ?? throw new Exception("WorksheetCommentsPart is missing.");
    var comments = commentsPart.Comments;
    var commentList = comments.GetFirstChild<CommentList>()
        ?? throw new Exception("CommentList is missing.");
    var commentElements = commentList.Elements<Comment>().ToList();
    if (commentElements.Count != 1)
        throw new Exception($"Expected 1 comment but got {commentElements.Count}.");

    var comment = commentElements[0];
    if (comment.Reference?.Value != "A1")
        throw new Exception($"Expected ref 'A1' but got '{comment.Reference?.Value}'.");

    var text = comment.GetFirstChild<CommentText>()
        ?? throw new Exception("CommentText is missing.");
    var runs = text.Elements<Run>().ToList();
    if (runs.Count != 2)
        throw new Exception($"Expected 2 runs but got {runs.Count}.");

    // Run 1: bold, sz 9, font Calibri
    var run1 = runs[0];
    var rpr1 = run1.GetFirstChild<RunProperties>()
        ?? throw new Exception("Run 1 RunProperties is missing.");
    if (rpr1.GetFirstChild<Bold>() == null)
        throw new Exception("Run 1 should have Bold.");
    var sz1 = rpr1.GetFirstChild<FontSize>();
    if (sz1 == null || sz1.Val?.Value != 9.0)
        throw new Exception($"Run 1 sz: expected 9 but got '{sz1?.Val?.Value}'.");
    var rfont1 = rpr1.GetFirstChild<RunFont>();
    if (rfont1 == null || rfont1.Val?.Value != "Calibri")
        throw new Exception($"Run 1 rFont: expected 'Calibri' but got '{rfont1?.Val?.Value}'.");
    var t1 = run1.GetFirstChild<Text>();
    if (t1?.Text != "Bold")
        throw new Exception($"Run 1 text: expected 'Bold' but got '{t1?.Text}'.");

    // Run 2: no formatting
    var run2 = runs[1];
    var t2 = run2.GetFirstChild<Text>();
    if (t2?.Text != " normal")
        throw new Exception($"Run 2 text: expected ' normal' but got '{t2?.Text}'.");
}
finally
{
    document.Dispose();
}
