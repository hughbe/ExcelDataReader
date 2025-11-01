using System.Globalization;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents a cell containing formula.
/// </summary>
internal sealed class XlsBiffFormulaCell : XlsBiffBlankCell
{
    // private FormulaFlags _flags;
    private readonly int _biffVersion;
    private bool _booleanValue;
    private CellError _errorValue;
    private double _xNumValue;
    private FormulaValueType _formulaType;
    private bool _initialized;

    internal XlsBiffFormulaCell(byte[] bytes, int biffVersion)
        : base(bytes)
    {
        _biffVersion = biffVersion;
    }

    [Flags]
    public enum FormulaFlags : ushort
    {
        AlwaysCalc = 0x0001,
        CalcOnLoad = 0x0002,
        SharedFormulaGroup = 0x0008
    }

    public enum FormulaValueType
    {
        Unknown,

        /// <summary>
        /// Indicates that a string value is stored in a String record that immediately follows this record. See[MS - XLS] 2.5.133 FormulaValue.
        /// </summary>
        String,

        /// <summary>
        /// Indecates that the formula value is an empty string.
        /// </summary>
        EmptyString,

        /// <summary>
        /// Indicates that the <see cref="BooleanValue"/> property is valid.
        /// </summary>
        Boolean,

        /// <summary>
        /// Indicates that the <see cref="ErrorValue"/> property is valid.
        /// </summary>
        Error,

        /// <summary>
        /// Indicates that the <see cref="XNumValue"/> property is valid.
        /// </summary>
        Number
    }

    public override bool IsEmpty => false;

    /// <summary>
    /// Gets the formula value type.
    /// </summary>
    public FormulaValueType FormulaType
    {
        get
        {
            LazyInit();
            return _formulaType;
        }
    }

    public bool BooleanValue
    {
        get
        {
            LazyInit();
            return _booleanValue;
        }
    }

    public CellError ErrorValue
    {
        get
        {
            LazyInit();
            return _errorValue;
        }
    }

    public double XNumValue
    {
        get
        {
            LazyInit();
            return _xNumValue;
        }
    }

    /*
    public FormulaFlags Flags
    {
        get
        {
            LazyInit();
            return _flags;
        }
    }
    */

    public string GetFormulaString(XlsFormulaReaderContext context)
    {
        // [MS-XLS] 2.5.198.3 CellParsedFormula
        // The CellParsedFormula structure specifies a formula (section 2.2.2) stored
        // in a cell.

        // For BIFF2, the rgce starts at offset 0x10.
        // For BIFF5 and later, the rgce starts at offset 0x14.
        int offset = _biffVersion < 5 ? 0x10 : 0x14;
        int cce;
        if (_biffVersion == 2)
        {
            // In BIFF2, cce is a single byte.
            cce = ReadByte(offset);
            offset += 1;
        }
        else
        {
            // cce (2 bytes): An unsigned integer that specifies the length of
            // rgce in bytes. MUST be greater than 0.
            cce = ReadUInt16(offset);
            offset += 2;
        }

        return XlsFormulaReader.ReadFormulaString(this, _biffVersion, offset, cce, offset + cce, context);
    }

    private void LazyInit()
    {
        if (_initialized)
            return;
        _initialized = true;

        if (_biffVersion == 2)
        {
            // _flags = (FormulaFlags)ReadUInt16(0xF);
            _xNumValue = ReadDouble(0x7);
            _formulaType = FormulaValueType.Number;
        }
        else
        {
            // _flags = (FormulaFlags)ReadUInt16(0xE);

            // fExprO (2 bytes): If fExprO is 0xFFFF, this structure specifies
            // a Boolean value, an error value, a string value, or a blank string value. If fExprO is not 0xFFFF, fExprO specifies the last two bytes of the Xnum.
            var formulaValueExprO = ReadUInt16(0xC);

            // byte1 (1 byte):  If fExprO is 0xFFFF, byte1 is an unsigned integer
            // that specifies the formula value type and MUST be a value from the
            // following table:
            if (formulaValueExprO != 0xFFFF)
            {
                _formulaType = FormulaValueType.Number;
                _xNumValue = ReadDouble(0x6);
            }
            else
            {
                // byte1 (1 byte):  If fExprO is 0xFFFF, byte1 is an unsigned
                // integer that specifies the formula value type and MUST be
                // a value from the following table:
                // 0x00
                // - String value. The string value is stored in a String
                // record that immediately follows this record.
                // 0x01
                // - Boolean value.
                // 0x02
                // - Error value.
                // 0x03
                // - Blank string value.
                // If fExprO is not 0xFFFF, byte1 specifies the first byte of the Xnum.
                var formulaValueByte1 = ReadByte(0x6);

                // byte3 (1 byte):  The meaning of byte3 is specified
                // in the following table:
                // fExprO is 0xFFFF and byte1 is 0x00
                // - byte3 is undefined and MUST be ignored.
                // fExprO is 0xFFFF and byte1 is 0x01
                // - byte3 specifies a Boolean value.
                // fExprO is 0xFFFF and byte1 is 0x02
                // - byte3 specifies a BErr.
                // fExprO is 0xFFFF and byte1 is 0x03
                // - byte3 is undefined and MUST be ignored.
                // fExprO is not 0xFFFF
                // - byte3 specifies the third byte of the Xnum.
                var formulaValueByte3 = ReadByte(0x8);
                switch (formulaValueByte1)
                {
                    case 0x00:
                        _formulaType = FormulaValueType.String;
                        break;
                    case 0x01:
                        _formulaType = FormulaValueType.Boolean;
                        _booleanValue = formulaValueByte3 != 0;
                        break;
                    case 0x02:
                        _formulaType = FormulaValueType.Error;
                        _errorValue = (CellError)formulaValueByte3;
                        break;
                    case 0x03:
                        _formulaType = FormulaValueType.EmptyString;
                        break;
                    default:
                        _formulaType = FormulaValueType.Unknown;
                        break;
                }
            }
        }
    }
}