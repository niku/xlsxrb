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

    var corePart = document.CoreFilePropertiesPart
        ?? throw new Exception("CoreFilePropertiesPart is missing.");

    var ns_dc = "http://purl.org/dc/elements/1.1/";
    var ns_cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

    var xml = System.Xml.Linq.XDocument.Load(corePart.GetStream());
    var root = xml.Root ?? throw new Exception("Root element missing.");

    Action<string, string, string> assertElement = (ns, name, expected) =>
    {
        var el = root.Element(System.Xml.Linq.XName.Get(name, ns));
        if (el == null)
            throw new Exception($"Element '{name}' not found in core properties.");
        if (el.Value != expected)
            throw new Exception($"Expected {name}='{expected}' but got '{el.Value}'.");
    };

    assertElement(ns_dc, "title", "My Title");
    assertElement(ns_dc, "subject", "My Subject");
    assertElement(ns_dc, "creator", "Alice");
    assertElement(ns_cp, "keywords", "ruby, xlsx");
    assertElement(ns_dc, "description", "A test document");
    assertElement(ns_cp, "lastModifiedBy", "Bob");
    assertElement(ns_cp, "revision", "3");
    assertElement(ns_cp, "category", "Reports");
    assertElement(ns_cp, "contentStatus", "Draft");
    assertElement(ns_dc, "language", "en-US");
}
finally
{
    document.Dispose();
}
