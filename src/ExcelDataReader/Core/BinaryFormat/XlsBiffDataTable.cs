using System.Globalization;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents a data table.
/// </summary>
internal sealed class XlsBiffDataTable : XlsBiffRecord
{
    private readonly int _biffVersion;
    private int _rowFirst;
    private int _rowLast;
    private int _colFirst;
    private int _colLast;
    private int _rowInputFirst;
    private int _colInputFirst;
    private int? _rowInputSecond;
    private int? _colInputSecond;

    internal XlsBiffDataTable(byte[] bytes, int biffVersion, BIFFRECORDTYPE recordType)
        : base(bytes)
    {
        _biffVersion = biffVersion;

        // [MS-XLS] 2.4.319 Table
        // The Table record specifies a data table (1). This record is preceded by a
        // single Formula record that defines the first cell in the data table (1).
        // Other Formula records that represent the rest of cells in the data table (1)
        // follow later in the file, not necessarily in a contiguous sequence. Formula
        // records that define the cells in the data table (1) MUST have the cell field
        // that is within the range specified in the ref field of this record and MUST
        // have their formula begin with PtgTbl. Also, each cell specified in the ref
        // field MUST have a Formula that is part of this table.
        int offset = 0;

        // ref (6 bytes):  A Ref structure that specifies the range of the data table (1).
        // The value of ref.rwFirst.rw MUST be greater than or equal to 1. The value of
        // ref.colFirst.col MUST be greater than or equal to 1.
        _rowFirst = ReadUInt16(offset);
        offset += 2;

        _rowLast = ReadUInt16(offset);
        offset += 2;

        _colFirst = ReadByte(offset);
        offset += 1;

        _colLast = ReadByte(offset);
        offset += 1;

        if (_biffVersion == 2)
        {
            // 0 = Do not recalculate the table, 1 = Always recalculate the table
            _ = ReadByte(offset);
            offset += 1;

            // 0 = Input data is in the column left of the table, formulas are in the row above the table
            // 1 = Input data is in the row above the table, formulas are in the column left of the table
            _ = ReadByte(offset);
            offset += 1;
        }
        else
        {
            // A - fAlwaysCalc (1 bit):  A bit that specifies whether this data table (1) is
            // recalculated as part of the next recalculation.
            // B - reserved1 (1 bit):  MUST be zero, and MUST be ignored.
            // C - fRw (1 bit):  A bit that specifies whether the input cell of a one-variable
            // data table is a row input cell or a column input cell. If the value is 1, the
            // input cell for a one-variable data table is a row input cell.
            // If the value of the fTbl2 field is 1, the value of fRw is undefined and MUST
            // be ignored.
            // D - fTbl2 (1 bit):  A bit that specifies whether the data table (1) is a
            // two-variable data table or a one-variable data table. If the value is 1, the
            // data table (1) is a two-variable data table.
            // E - fDeleted1 (1 bit):  A bit that specifies whether the cell referenced in the input
            // cell specified by the rwInpRw and colInpRw fields is deleted.
            // F - fDeleted2 (1 bit):  A bit that specifies whether the cell referenced in the input
            // cell specified by the rwInpCol and colInpCol fields is deleted.
            // reserved2 (10 bits):  MUST be zero, and MUST be ignored.
            _ = ReadUInt16(offset);
            offset += 2;
        }

        // rwInpRw (2 bytes):  A RwU structure that specifies either the row of a row input
        // cell or the row of a column input cell. If the value of the fTbl2 field is 0 and
        // the value of the fRw field is 0, the value of rwInpRw specifies the row of a
        // column input cell; for any other combination of the fTbl2 and fRw fields,
        // rwInpRw specifies the row of a row input cell. If the value of fDeleted
        // field is 1, the value of rwInpRw MUST be 65535.
        // If fTbl2 is 1, the following statement (1) holds.
        // If fTbl2 is 0, exactly one of these statements holds:
        // rwInpRw and colInpRw MUST specify a cell outside the bounds specified by
        // ref.rwFirst – 1, ref.rwLast, ref.colFirst – 1, and ref.colLast.
        // rwInpRw and colInpRw MUST be equal to ref.rwFirst -1 and ref.colFirst – 1,
        // respectively.
        _rowInputFirst = ReadUInt16(offset);
        offset += 2;

        // colInpRw (2 bytes):  A Col_NegativeOne structure that specifies either the
        // column of a row input cell or the column of a column input cell. If the
        // value of the fTbl2 field is 0 and the value of fRw field is 0, the value
        // of colInpRw specifies the column of the column input cell; for any other
        // combination of the fTbl2 and fRw fields, colInpRw specifies the column of a
        // row input cell. If the value of the fDeleted field is 1, the value of colInpRw
        // MUST be -1. If the value of the fDeleted field is 0, the value of colInpRw
        // MUST be greater than or equal to 0.
        _colInputFirst = ReadUInt16(offset);
        offset += 2;

        if (_biffVersion > 2 || recordType == BIFFRECORDTYPE.DATATABLE2)
        {
            // rwInpCol (2 bytes):  A RwU structure that specifies the row of the column
            // input cell. The restrictions on the value of rwInpCol are dictated by the
            // value of the fTbl2 field and the value of the fDeleted2 field, as
            // specified in the following table:
            // fTbl2 fDeleted2 rwInpCol
            // 1 1 The value MUST be 65535.
            // 1 0 If the colInpCol is a value between ref.colFirst – 1 and ref.colLast inclusive, rwInpCol MUST NOT be a value between ref.rwFirst – 1 and ref.rwLast inclusive.
            // 0 1 or 0 Undefined and MUST be ignored.
            _rowInputSecond = ReadUInt16(offset);
            offset += 2;

            // colInpCol (2 bytes):  A Col_NegativeOne structure that specifies the column
            // of the column input cell. The restrictions on the value of colInpCol are
            // dictated by the value of the fTbl2 field and the value of the fDeleted2
            // field, as specified in the following table:
            // fTbl2 fDeleted2 colInpCol
            // 1 1 The value MUST be -1.
            // 1 0 The value MUST be greater than or equal to 0. If the rwInpCol is a value
            // between ref.rwFirst – 1 and ref.rwLast inclusive, colInpCol MUST NOT be a
            // value between ref.colFirst – 1 and ref.colLast, inclusive.
            // 0 1 or 0 Undefined and MUST be ignored.
            _colInputSecond = ReadUInt16(offset);
            offset += 2;
        }
    }

    public int RowFirst => _rowFirst;

    public int RowLast => _rowLast;

    public int ColFirst => _colFirst;

    public int ColLast => _colLast;

    public int RowInputFirst => _rowInputFirst;

    public int ColInputFirst => _colInputFirst;
    
    public int? RowInputSecond => _rowInputSecond;

    public int? ColInputSecond => _colInputSecond;
}
