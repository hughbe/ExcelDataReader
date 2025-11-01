namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents a shared formula.
/// </summary>
internal sealed class XlsBiffSharedFormula : XlsBiffRecord
{
    private readonly int _biffVersion;
    private readonly int _rowFirst;
    private readonly int _rowLast;
    private readonly int _colFirst;
    private readonly int _colLast;
    private readonly int _formulaLength;
    private readonly int _formulaOffset;

    internal XlsBiffSharedFormula(byte[] bytes, int biffVersion)
        : base(bytes)
    {
        _biffVersion = biffVersion;

        if (_biffVersion < 5)
        {
            throw new NotSupportedException("Shared formulas are not supported in BIFF2-5.");
        }

        // [MS-XLS] 2.4.260 ShrFmla
        // The ShrFmla record specifies a formula (section 2.2.2) that is shared across
        // multiple cells.  This record specifies a file size optimization.  It is used
        // with the Formula record to compress the amount of storage required for the
        // formula. This record is preceded by a single Formula record that specifies
        // the first cell in the range that uses this shared formula.  Other Formula
        // records that use this shared formula follow later in the file, not necessarily
        // in a contiguous sequence.  Formula records that use this shared formula have
        // the Formula.fShrFmla bit set, and a Formula.cell that is within the range
        // specified in the ref field of this record.
        int offset = 0;

        // ref (6 bytes): A RefU structure that specifies the range of cells that use
        // this shared formula.  Cells in this range do not have to use the shared formula.
        _rowFirst = ReadUInt16(offset);
        offset += 2;

        _rowLast = ReadUInt16(offset);
        offset += 2;

        _colFirst = ReadByte(offset);
        offset += 1;

        _colLast = ReadByte(offset);
        offset += 1;

        // reserved (8 bits): MUST be zero, and MUST be ignored.
        _ = ReadByte(offset);
        offset += 1;

        // cUse (8 bits): An unsigned integer that specifies the number of cells that
        // use this shared formula.
        _ = ReadByte(offset);
        offset += 1;

        // formula (variable): A SharedParsedFormula structure that specifies the shared
        // formula.
        int cce = ReadUInt16(offset);
        offset += 2;

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