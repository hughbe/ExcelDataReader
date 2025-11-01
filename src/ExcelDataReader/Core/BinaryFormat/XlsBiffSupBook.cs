using System.Globalization;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Specifies an external workbook.
/// </summary>
internal sealed class XlsBiffSupBook : XlsBiffRecord
{
    private readonly XlsBiffSupBookType _type;
    private readonly string _url;
    private readonly List<string> _sheets;

    internal XlsBiffSupBook(byte[] bytes, int biffVersion)
        : base(bytes)
    {
        int offset = 0;

        // ctab (2 bytes):  An undefined field, a reserved field, or an unsigned integer
        // that specifies the number of sheets in a referenced external workbook. The
        // type and meaning of this field is dependent on the type of supporting link
        // specified by the cch and virtPath fields, and is defined in the following table:
        // Self-referencing
        // - Undefined and MUST be ignored.
        // Same-sheet referencing DDE data source referencing OLE data source referencing
        // - Reserved. MUST be 0x0000.
        // Add-in referencing
        // - Reserved. MUST be 0x0001.
        // External workbook referencing
        // - An unsigned integer that specifies the count of sheets in the referenced
        // external workbook.
        // Unused
        // An unsigned integer that specifies the count of sheets in the external workbook formerly referenced by this supporting link, if this supporting link was an external workbook referencing type, when used. Otherwise, this value MUST be 0x0000.
        int numberOfSheets = ReadUInt16(offset);
        offset += 2;

        // cch (2 bytes): An unsigned integer that specifies a type of supporting link
        // or specifies the length of the string in virtPath. MUST be a value from the
        // following table:
        // 0x0401 This record specifies a self-referencing supporting link.
        // Peek at the first two bytes to determine if this is an external reference,
        // internal reference, add-in function or DDE/OLE link.
        // 0x3A01 This record specifies an add-in referencing type of supporting link.
        // The names of all add-in functions implemented by XLL, or COM automation add-ins
        // that are referenced by formulas in this workbook, MUST be specified in the
        // ExternName records that follow this record.
        // 0x0001 to 0x00ff (inclusive) The type of supporting link specified by this
        // record is specified by virtPath.  This value is the count of characters in
        // virtPath.
        int cch = ReadUInt16(offset);
        if (cch == 0x0401)
        {
            // Internal 3D reference.
            // No URL or sheet Urls are stored.
            _type = XlsBiffSupBookType.Internal3DReference;
            return;
        }
        else if (cch == 0x3A01)
        {
            // Add-in function or DDE/OLE link.
            // No URL or sheet Urls are stored.
            _type = XlsBiffSupBookType.AddInFunctionOrDdeOleLink;
            return;
        }
        else
        {
            _type = XlsBiffSupBookType.ExternalWorkbookReference;
        }   
        
            // virtPath (variable):  An XLUnicodeStringNoCch structure that specifies the type
        // of supporting link and, if applicable, the target of that supporting link.
        // This field MUST exist if and only if the value of cch is between 0x0001 and
        // 0x00ff (inclusive). The length of the string in this field MUST be equal to cch.
        // The contents of this field MUST be a value from the following table:
        _url = XlsBiffEncodedUrl.Decode(this, biffVersion, offset, out int bytesRead, out _);
        offset += bytesRead;

        // rgst (variable):  An array of XLUnicodeString structures that specify sheet
        // names in the external workbook. This field MUST exist if and only if the
        // supporting link type specified by cch and virtPath is external workbook
        // referencing or unused. If this field exists, the number of elements in
        // this array MUST be equal to ctab. The contents and meaning of this array
        // are defined in the following table:
        var sheets = new List<string>(numberOfSheets);
        for (int i = 0; i < numberOfSheets; i++)
        {
            string sheetName = ReadXLUnicodeString(offset, out int sheetNameBytesRead);
            offset += sheetNameBytesRead;
            sheets.Add(sheetName);
        }

        _sheets = sheets;
    }

    public string Url => _url;

    public XlsBiffSupBookType Type => _type;

    public List<string> Sheets => _sheets;
}
