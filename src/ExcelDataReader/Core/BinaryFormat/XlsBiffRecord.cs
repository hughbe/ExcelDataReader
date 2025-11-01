using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents basic BIFF record.
/// Base class for all BIFF record types.
/// </summary>
internal class XlsBiffRecord
{
    protected const int ContentOffset = 4;
    
    public XlsBiffRecord(byte[] bytes)
    {
        if (bytes.Length < 4)
            throw new ArgumentException(Errors.ErrorBiffRecordSize);
        Bytes = bytes;
    }

    /// <summary>
    /// Gets the type Id of this entry.
    /// </summary>
    public BIFFRECORDTYPE Id => (BIFFRECORDTYPE)BitConverter.ToUInt16(Bytes, 0);

    /// <summary>
    /// Gets the data size of this entry.
    /// </summary>
    public ushort RecordSize => BitConverter.ToUInt16(Bytes, 2);

    /// <summary>
    /// Gets the whole size of structure.
    /// </summary>
    public int Size => ContentOffset + RecordSize;
    
    internal byte[] Bytes { get; }

    #if NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER
    public virtual void Return()
    {        
        System.Buffers.ArrayPool<byte>.Shared.Return(Bytes);
    }
    #endif

    public byte ReadByte(int offset)
    {
        return Buffer.GetByte(Bytes, ContentOffset + offset);
    }

    public ushort ReadUInt16(int offset)
    {
        return BitConverter.ToUInt16(Bytes, ContentOffset + offset);
    }

    public uint ReadUInt32(int offset)
    {
        return BitConverter.ToUInt32(Bytes, ContentOffset + offset);
    }

    public ulong ReadUInt64(int offset)
    {
        return BitConverter.ToUInt64(Bytes, ContentOffset + offset);
    }

    public short ReadInt16(int offset)
    {
        return BitConverter.ToInt16(Bytes, ContentOffset + offset);
    }

    public int ReadInt32(int offset)
    {
        return BitConverter.ToInt32(Bytes, ContentOffset + offset);
    }

    public long ReadInt64(int offset)
    {
        return BitConverter.ToInt64(Bytes, ContentOffset + offset);
    }

    public byte[] ReadArray(int offset, int size)
    {
        byte[] tmp = new byte[size];
        Buffer.BlockCopy(Bytes, ContentOffset + offset, tmp, 0, size);
        return tmp;
    }

    public float ReadFloat(int offset)
    {
        return BitConverter.ToSingle(Bytes, ContentOffset + offset);
    }

    public double ReadDouble(int offset)
    {
        return BitConverter.ToDouble(Bytes, ContentOffset + offset);
    }

    public string ReadByteString(int offset, out int bytesRead)
    {
        // All Excel file formats up to BIFF5 contain simple byte strings.
        // The byte string consists of the length of the string followed by the character array.
        // The length is stored either as 8-bit value or as 16-bit value, depending on the
        // current record. The string is not zero-terminated.
        byte cch = ReadByte(offset);
        bytesRead = 1 + cch;

        return ReadAsciiString(offset + 1, cch);
    }

    public string ReadAsciiString(int offset, int length)
    {
        return Encoding.ASCII.GetString(
            Bytes,
            ContentOffset + offset,
            length);
    }

    public string ReadShortXLUnicodeString(int offset, out int bytesRead)
    {
        // [MS-XLS] 2.5.240 ShortXLUnicodeString
        // The ShortXLUnicodeString structure specifies a Unicode string.

        // cch (1 bytes): An unsigned integer that specifies
        // the count of characters in the string.
        byte cch = ReadByte(offset);
        offset += 1;

        // A - fHighByte (1 bit): A bit that specifies whether the characters in rgb
        // are double-byte characters. MUST be a value from the following table:
        // 0x0 All the characters in the string have a high byte of 0x00 and only the
        // low bytes are in rgb.
        // 0x1 All the characters in the string are saved as double-byte characters in
        // rgb.
        // reserved (7 bits): MUST be zero, and MUST be ignored.
        // rgb (variable): An array of bytes that specifies the characters. If fHighByte
        // is 0x0, the size of the array MUST be equal to the value of cch. If fHighByte
        // is 0x1, the size of the array MUST be equal to the value of cch*2.
        string result = ReadXLUnicodeStringNoCch(offset, cch, out var stringBytesRead);
        bytesRead = 1 + stringBytesRead;
        return result;
    }

    public string ReadXLUnicodeString(int offset, out int bytesRead)
    {
        // [MS-XLS] 2.5.294 XLUnicodeString
        // The XLUnicodeString structure specifies a Unicode string.
        bytesRead = 0;

        // cch (2 bytes): An unsigned integer that specifies the count of characters in
        // the string.
        ushort cch = ReadUInt16(offset + bytesRead);
        bytesRead += 2;

        // A - fHighByte (1 bit): A bit that specifies whether the characters in rgb
        // are double-byte characters. MUST be a value from the following table:
        // 0x0 All the characters in the string have a high byte of 0x00 and only the
        // low bytes are in rgb.
        // 0x1 All the characters in the string are saved as double-byte characters in
        // rgb.
        // reserved (7 bits): MUST be zero, and MUST be ignored.
        // rgb (variable): An array of bytes that specifies the characters. If fHighByte
        // is 0x0, the size of the array MUST be equal to the value of cch. If fHighByte
        // is 0x1, the size of the array MUST be equal to the value of cch*2.
        string result = ReadXLUnicodeStringNoCch(offset + bytesRead, cch, out var stringBytesRead);
        bytesRead += stringBytesRead;
        return result;
    }

    public string ReadXLUnicodeStringNoCch(int offset, int cch, out int bytesRead)
    {
        bytesRead = 0;

        // [MS-XLS] 2.5.296 XLUnicodeStringNoCch
        // The XLUnicodeStringNoCch structure specifies a Unicode string.
        // When an XLUnicodeStringNoCch is used, the count of characters in the
        // string MUST be specified in the structure that uses the XLUnicodeStringNoCch.
        // A - fHighByte (1 bit): A bit that specifies whether the characters in rgb
        // are double-byte characters. MUST be a value from the following table:
        // 0x0 All the characters in the string have a high byte of 0x00 and only
        // the low bytes are in rgb.
        // 0x00 All the characters in the string are saved as double-byte characters
        // in rgb.
        // reserved (7 bits): MUST be zero, and MUST be ignored.
        byte optionFlags = ReadByte(offset + bytesRead);
        bytesRead += 1;

        bool isDoubleByte = (optionFlags & 0x01) != 0;

        // rgb (variable): An array of bytes that specifies the characters.
        // If fHighByte is 0x0, the size of the array MUST be equal to the count of
        // characters in the string. If fHighByte is 0x1, the size of the array MUST
        // be equal to 2 times the count of characters in the string.
        if (isDoubleByte)
        {
            var result = ReadUnicodeStringNoCch(offset + bytesRead, cch, out var stringBytesRead);
            bytesRead += stringBytesRead;
            return result;
        }
        else
        {
            bytesRead += cch;
            return Encoding.ASCII.GetString(Bytes, ContentOffset + offset, cch);
        }
    }

    public string ReadUnicodeStringNoCch(int offset, int cch, out int bytesRead)
    {
        bytesRead = cch * 2;
        return Encoding.Unicode.GetString(Bytes, ContentOffset + offset, cch * 2);
    }
}
