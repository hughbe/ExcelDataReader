namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// [MS-XLS] 2.5.344 XTI
/// The XTI structure specifies a supporting link and scope information about that
/// supporting link.
/// </summary>
internal sealed class XlsBiffXti
{
    private readonly int _workbookIndex;
    private readonly int _firstSheetIndex;
    private readonly int _lastSheetIndex;

    internal XlsBiffXti(byte[] bytes, int offset)
    {
        // iSupBook (2 bytes): An unsigned integer that specifies the zero-based index of
        // a SupBook record in the collection of SupBook records in the Globals Substream
        // ABNF. The referenced SupBook specifies the supporting link referenced by this
        // structure. This value MUST be less than the number of SupBook records in this
        // file.
        _workbookIndex = BitConverter.ToUInt16(bytes, offset);
        offset += 2;

        // itabFirst (2 bytes): A signed integer that specifies the scope of the supporting
        // link, and if a scope is specified, the first sheet in the scope of that
        // supporting link. If the type of supporting link specified by the cch and
        // virtPath fields of the SupBook record is same-sheet referencing, add-in
        // referencing, DDE data source referencing, or OLE data source referencing,
        // then no scope is specified and this value MUST be -2. Otherwise, this field
        // MUST contain a value from the following table:
        // -2 Workbook-level reference that applies to the entire workbook.
        // -1 Sheet-level reference. The first sheet in the reference could not be found.
        // >= 0 Sheet-level reference. This specifies the first sheet in the reference.
        // If the supporting link type is unused or external workbook referencing, then
        // this value specifies the zero-based index of an XLUnicodeString in the rgst
        // field of the SupBook record specified in iSupBook. This XLUnicodeString
        // specifies the name of the first sheet within the external workbook that is
        // in scope. This sheet MUST be a worksheet or macro sheet.
        // If the supporting link type is self-referencing, then this value specifies
        // the zero-based index of a BoundSheet8 record in the Globals Substream ABNF
        // that specifies the first sheet within the scope of this reference. This
        // sheet MUST be a worksheet or a macro sheet.
        _firstSheetIndex = BitConverter.ToInt16(bytes, offset);
        offset += 2;

        // itabLast (2 bytes): A signed integer that specifies the scope of the supporting
        // link, and if a scope is specified, the last sheet in the scope of that
        // supporting link. If the type of supporting link specified by the cch and
        // virtPath fields of the SupBook record is same-sheet referencing, add-in
        // referencing, DDE data source referencing, or OLE data source referencing,
        // then no scope is specified and this value MUST be -2. Otherwise, this field
        // MUST contain a value from the following table:
        // -2 Workbook-level reference that applies to the entire workbook. MUST NOT be
        // used if itabFirst is not equal to -2.
        // -1 Sheet-level reference. The last sheet in the reference could not be found.
        // SHOULD NOT<192> be used if itabFirst is equal to -2.
        // >= 0 Sheet-level reference. This specifies the last sheet in the reference. MUST
        // NOT be used if itabFirst is equal to -2.
        // If the supporting link type is unused or referring to an external workbook, then
        // this value specifies the zero-based index of an XLUnicodeString in the rgst
        // field of the SupBook record specified in iSupBook. This XLUnicodeString
        // specifies the name of the last sheet within the external workbook that is
        // in scope. This sheet MUST be a worksheet or macro sheet.
        // If the supporting link type is self-referencing, then this value specifies
        // the zero-based index of a BoundSheet8 record in the Globals Substream ABNF
        // that specifies the last sheet within the scope of this reference. This sheet
        // MUST be a worksheet or a macro sheet.
        _lastSheetIndex = BitConverter.ToInt16(bytes, offset);
        offset += 2;
    }

    public int WorkbookIndex => _workbookIndex;

    public int FirstSheetIndex => _firstSheetIndex;

    public int LastSheetIndex => _lastSheetIndex;
}
