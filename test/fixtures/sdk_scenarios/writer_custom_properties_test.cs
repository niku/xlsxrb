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

    var customPropsPart = document.CustomFilePropertiesPart;
    if (customPropsPart == null)
    {
        throw new Exception("CustomFilePropertiesPart is missing.");
    }

    var props = customPropsPart.Properties;
    if (props == null)
    {
        throw new Exception("Properties element is missing.");
    }

    var propList = props.Elements<CustomDocumentProperty>().ToList();
    if (propList.Count < 3)
    {
        throw new Exception($"Expected at least 3 custom properties but got {propList.Count}.");
    }

    var project = propList.FirstOrDefault(p => p.Name?.Value == "Project");
    if (project == null)
    {
        throw new Exception("Custom property 'Project' not found.");
    }
    var projectVal = project.VTLPWSTR?.Text;
    if (projectVal != "Alpha")
    {
        throw new Exception($"Expected Project='Alpha' but got '{projectVal}'.");
    }

    var version = propList.FirstOrDefault(p => p.Name?.Value == "Version");
    if (version == null)
    {
        throw new Exception("Custom property 'Version' not found.");
    }
    var versionVal = version.VTInt32?.Text;
    if (versionVal != "42")
    {
        throw new Exception($"Expected Version=42 but got '{versionVal}'.");
    }

    var active = propList.FirstOrDefault(p => p.Name?.Value == "Active");
    if (active == null)
    {
        throw new Exception("Custom property 'Active' not found.");
    }
    var activeVal = active.VTBool?.Text;
    if (activeVal != "true")
    {
        throw new Exception($"Expected Active=true but got '{activeVal}'.");
    }
}
finally
{
    document.Dispose();
}
