namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents an array formula.
/// </summary>
internal sealed class XlsBiffArray : XlsBiffRecord
{
    private readonly int _biffVersion;
    private readonly int _rowFirst;
    private readonly int _rowLast;
    private readonly int _colFirst;
    private readonly int _colLast;
    private readonly int _formulaLength;
    private readonly int _formulaOffset;

    internal XlsBiffArray(byte[] bytes, int biffVersion)
        : base(bytes)
    {
        _biffVersion = biffVersion;

        // [MS-XLS] 2.4.45 Array
        // The Array record specifies an array formula (section 2.2.2) for a range of cells
        // that performs calculations on one or more sets of values, and then returns either
        // a single result or multiple results across a continuous range of cells. This record
        // is preceded by a single Formula record (section 2.4.127) that defines the first
        // cell in the range that uses this array formula (section 2.2.2). Other Formula
        // records (section 2.4.127) that use this array formula (section 2.2.2) follow
        // later in the file, not necessarily in a contiguous sequence. Formula records
        // (section 2.4.127) that use this array formula (section 2.2.2) MUST have a cell
        // field that is within the range specified in the ref field of this record and
        // MUST have their formula begin with PtgExp (section 2.5.198.58). Also, each
        // cell specified in the ref field MUST have a Formula (section 2.4.127) that
        // uses this array formula (section 2.2.2).
        int offset = 0;

        // ref (6 bytes): A Ref structure (section 2.5.207) that specifies the range of
        // the array formula (section 2.2.2).
        _rowFirst = ReadUInt16(offset);
        offset += 2;

        _rowLast = ReadUInt16(offset);
        offset += 2;

        _colFirst = ReadByte(offset);
        offset += 1;

        _colLast = ReadByte(offset);
        offset += 1;

        if (biffVersion == 2)
        {
            // BIFF2 Array structure is different.
            _ = ReadByte(offset);
            offset += 1;
        }
        else
        {
            // A - fAlwaysCalc (1 bit): A bit that specifies whether the array formula
            // (section 2.2.2) needs to be calculated during the next recalculation.
            // reserved (15 bits): MUST be zero, and MUST be ignored.
            _ = ReadUInt16(offset);
            offset += 2;
        }

        if (_biffVersion >= 5)
        {
            // unused (4 bytes): Undefined and MUST be ignored.
            _ = ReadUInt32(offset);
            offset += 4;
        }

        // formula (variable): An ArrayParsedFormula structure (section 2.5.198.1) that
        // specifies the array formula (section 2.2.2).
        int cce;
        if (_biffVersion == 2)
        {
            // BIFF2 uses a single byte for cce.
            cce = ReadByte(offset);
            offset += 1;
        }
        else
        {
            // Size of the following formula data (cce)
            cce = ReadUInt16(offset);
            offset += 2;
        }

        _formulaOffset = offset;
        _formulaLength = cce;
    }

    public int RowFirst => _rowFirst;

    public int RowLast => _rowLast;

    public int ColFirst => _colFirst;

    public int ColLast => _colLast;

    public string GetFormulaString(XlsFormulaReaderContext context)
    {
        return XlsFormulaReader.ReadFormulaString(this, _biffVersion, _formulaOffset, _formulaLength, _formulaOffset + _formulaLength, context);
    }
}