namespace ExcelReader.Core.Writer.Internal
{
    internal static class XlsxConstants
    {
        internal const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        internal const string RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        internal const string PackageRelationshipsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        internal const string ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";

        internal const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        internal const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        internal const string StylesContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        internal const string RelationshipsContentType = "application/vnd.openxmlformats-package.relationships+xml";

        internal const string WorkbookRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        internal const string WorksheetRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        internal const string StylesRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        internal const string SharedStringsRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    }
}
