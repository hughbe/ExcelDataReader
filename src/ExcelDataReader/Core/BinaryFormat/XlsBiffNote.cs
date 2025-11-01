using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents a NOTE record.
/// </summary>
internal sealed class XlsBiffNote : XlsBiffRecord
{
    private readonly int _biffVersion;
    private int _row;
    private int _col;
    private int _length;

    internal XlsBiffNote(byte[] bytes, int biffVersion)
        : base(bytes)
    {
        _biffVersion = biffVersion;

        // [MS-XLS] 2.5.186 NoteSh
        // The NoteSh structure specifies a comment associated with a cell.
        int offset = 0;

        // row (2 bytes): A RW that specifies the row of the cell to which this comment is
        // associated.
        _row = ReadUInt16(offset);
        offset += 2;

        // col (2 bytes): A Col that specifies the column of the cell to which this comment
        // is associated.
        _col = ReadUInt16(offset);
        offset += 2;

        if (_biffVersion <= 5)
        {
            _length = ReadUInt16(offset);
            offset += 2;
        }
        else
        {
            // A - reserved1 (1 bit): MUST be zero and MUST be ignored.
            // B - fShow (1 bit): A bit that specifies whether the comment is shown at all times.
            // C - reserved2 (1 bit): MUST be zero and MUST be ignored.
            // D - unused1 (1 bit): Undefined and MUST be ignored.
            // E - reserved3 (3 bits): MUST be zero and MUST be ignored.
            // F - fRwHidden (1 bit): A bit that specifies whether the row specified by row is hidden.
            // G - fColHidden (1 bit): A bit that specifies whether the column specified by col is hidden.
            //  reserved4 (7 bits): MUST be zero and MUST be ignored.
            _ = ReadUInt16(offset);
            offset += 2;

            // idObj (2 bytes): An ObjId that specifies the Obj record that specifies the
            // comment text.
            _ = ReadUInt16(offset);
            offset += 2;

            // stAuthor (variable): An XLUnicodeString that specifies the name of the comment
            // author. String length MUST be greater than or equal to 1 and less than or equal
            // to 54.
            _ = ReadXLUnicodeString(offset, out var bytesRead);
            offset += bytesRead;

            // unused2 (1 byte): Undefined and MUST be ignored.
            _ = ReadByte(offset);
            offset += 1;
        }
    }

    public int RowIndex => _row;

    public int ColumnIndex => _col;

    public int Length => _length;

    public string GetText(Encoding encoding)
    {
        if (_biffVersion <= 5)
        {
            return encoding.GetString(Bytes, ContentOffset + 6, Math.Min(2048, _length));
        }
        
        return string.Empty;
    }
}
