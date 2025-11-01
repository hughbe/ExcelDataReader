using System.Globalization;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents a defined name.
/// </summary>
internal sealed class XlsBiffDefinedName : XlsBiffRecord
{
    private readonly string _name;
    private readonly int _formulaOffset;
    private readonly int _formulaLength;
    private readonly int _formulaExtraOffset;

    internal XlsBiffDefinedName(byte[] bytes, int biffVersion)
        : base(bytes)
    {
        int offset = 0;
        ushort cce;

        // [MS_XLS] 2.4.150 Lbl
        // The Lbl record specifies a defined name.
        if (biffVersion == 2)
        {
            // Option flags:
            // Bit Mask Contents
            // 1 02H 1 = Function macro or command macro
            // 2 04H 1 = Complex function (array formula or user defined)
            _ = ReadByte(offset);
            offset += 1;

            // If name is function macro or command macro
            // (see option flags above):
            // 01H = Function macro, 02H = Command macro
            _ = ReadByte(offset);
            offset += 1;
        }
        else
        {
            // A - fHidden (1 bit): A bit that specifies whether the defined name is not
            // visible in the list of defined names.
            // B - fFunc (1 bit): A bit that specifies whether the defined name represents
            // an Excel macro (XLM). If this bit is 1, fProc MUST also be 1.
            // C - fOB (1 bit): A bit that specifies whether the defined name represents
            // a Visual Basic for Applications (VBA) macro. If this bit is 1, the fProc
            // MUST also be 1.
            // D - fProc (1 bit): A bit that specifies whether the defined name represents
            // a macro.
            // E - fCalcExp (1 bit): A bit that specifies whether rgce contains a call to
            // a function that can return an array.
            // F - fBuiltin (1 bit): A bit that specifies whether the defined name
            // represents a built-in name.
            // fGrp (6 bits): An unsigned integer that specifies the function category
            // for the defined name. MUST be less than or equal to 31. The values 17
            // to 31 are user-defined. User-defined values are specified in the
            // FnGroupName record. The values 0 to 16 are defined as specified in
            // the following table:
            // G - reserved1 (1 bit): MUST be zero, and MUST be ignored.
            // H - fPublished (1 bit): A bit that specifies whether the defined name is
            // published. This bit is ignored if the fPublishedBookItems field of the
            // BookExt_Conditional12 structure is 0.
            // I - fWorkbookParam (1 bit): A bit that specifies whether the defined name
            // is a workbook parameter.
            // J - reserved2 (1 bit): MUST be zero, and MUST be ignored.
            _ = ReadUInt16(offset);
            offset += 2;
        }

        // chKey (1 byte):  The unsigned integer value of the ASCII character that
        // specifies the shortcut key for the macro represented by the defined name.
        // MUST be 0 (no shortcut key) if fFunc is 1 or if fProc is 0. Otherwise
        // MUST be greater than or equal to 0x41 and less than or equal to 0x5A,
        // or greater than or equal to 0x61 and less than or equal to 0x7A.
        _ = ReadByte(offset);
        offset += 1;

        // cch (1 byte): An unsigned integer that specifies the number of characters
        // in Name. MUST be greater than or equal to zero.
        var cch = ReadByte(offset);
        offset += 1;

        // cce (2 bytes): An unsigned integer that specifies length of rgce in bytes.
        if (biffVersion == 2)
        {
            // BIFF2 uses a single byte for cce.
            cce = ReadByte(offset);
            offset += 1;
        }
        else
        {
            cce = ReadUInt16(offset);
            offset += 2;
        }

        if (biffVersion >= 5)
        {
            // reserved3 (2 bytes): MUST be zero, and MUST be ignored.
            _ = ReadUInt16(offset);
            offset += 2;

            // itab (2 bytes): An unsigned integer that specifies if the defined name is
            // a local name, and if so, which sheet it is on. If itab is not 0, the
            // defined name is a local name and the value MUST be a one-based index to
            // the collection of BoundSheet8 records as they appear in the Globals
            // Substream.
            _ = ReadUInt16(offset);
            offset += 2;

            // reserved4 (1 byte): MUST be zero, and MUST be ignored.
            _ = ReadByte(offset);
            offset += 1;

            // reserved5 (1 byte): MUST be zero, and MUST be ignored.
            _ = ReadByte(offset);
            offset += 1;

            // reserved6 (1 byte): MUST be zero, and MUST be ignored.
            _ = ReadByte(offset);
            offset += 1;

            // reserved7 (1 byte): MUST be zero, and MUST be ignored.
            _ = ReadByte(offset);
            offset += 1;
        }

        // Name (variable): An XLUnicodeStringNoCch structure that specifies the
        // name for the defined name. If fBuiltin is 0, this field MUST satisfy
        // the same restrictions as the name field of the XLNameUnicodeString
        // structure. If fBuiltin is 1, this field is for a built-in name. Each
        // built-in name has a zero-based index value associated with it. A
        // built-in name or its index value MUST be used for this field. The
        // built-in names are defined in the following table:
        if (biffVersion <= 5)
        {
            _name = ReadAsciiString(offset, cch);
            offset += cch;
        }
        else
        {
            _name = ReadXLUnicodeStringNoCch(offset, cch, out int nameBytesRead);
            offset += nameBytesRead;
        }

        // rgce (variable): A NameParsedFormula structure that specifies the formula for
        // the defined name.
        _formulaOffset = offset;
        _formulaLength = cce;
        offset += cce;

        if (biffVersion == 2)
        {
            // Duplicate of the formula data size field (sz)
            _ = ReadByte(offset);
            offset += 1;
        }

        // rgbExtra (variable): Extra data for the formula follows the rgce and the sz field.
        _formulaExtraOffset = offset;
    }

    public string Name => _name;

    public int FormulaOffset => _formulaOffset;

    public int FormulaLength => _formulaLength;

    public int FormulaExtraOffset => _formulaExtraOffset;
}
