using System.Globalization;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Represents an external name.
/// </summary>
internal sealed class XlsBiffExternalName : XlsBiffRecord
{
    private readonly int? _externalSheetIndex;
    private readonly string _name;
    private readonly int _dataOffset;
    private readonly int _dataLength;

    internal XlsBiffExternalName(byte[] bytes, int biffVersion)
        : base(bytes)
    {
        int offset = 0;

        // [MS_XLS] 2.4.105 ExternName
        // The ExternName record specifies an external defined name, a User Defined Function
        // (UDF) reference on a XLL or COM add-in, a DDE data item or an OLE data item,
        // depending on the value of the virtPath field in the preceding SupBook record.
        // If the cch field in the preceding SupBook record is 0x3A01, then this record
        // specifies a UDF reference. Otherwise if the virtPath field in the preceding
        // SupBook record conforms to the ole-link rule specified in the VirtualPath ABNF,
        // then this record specifies a DDE data item or an OLE data item. Otherwise, this
        // record specifies an external defined name.
        if (biffVersion == 2)
        {
            // BIFF2 ExternName structure is different.
            _name = ReadByteString(offset, out int nameBytesRead);
            offset += nameBytesRead;

            // Open Office specifies that there is formula data here,
            // but this is incorrect.
            // Microsoft documentation from 1988 states that 
            // "When the externally referenced name is a DDE topic, Excel may
            // append the most recent values for the topic to the EXTERNNAME
            // record.  The values are written in the same format as array
            // constant values in parsed expressions.  See the explanation
            // of "ptgArray" in the "Operand Tokens - Base" section for a
            // full description of this format.
            // https://github.com/xieguigang/sciBASIC/blob/master/mime/application%25vnd.openxmlformats-officedocument.spreadsheetml.sheet/Excel/XLS/BIFF/excel.txt
            // This is 6 bytes per name, but it can be ignored for our purposes.
            _dataOffset = offset;
            _dataLength = 0;
        }
        else if (biffVersion <= 8)
        {
            // Option flags (see below)
            // 0 0001H 0 = Standard name; 1 = Built-in name
            // 1 0002H 0 = Manual link; 1 = Automatic link (DDE links and OLE links only)
            // 2 0004H 1 = Picture link (DDE links and OLE links only)
            // 3 0008H 1 = This is the “StdDocumentName” identifier (DDE links only)
            // 4 0010H 1 = OLE link
            // 14-5 7FE0H Clipboard format of last successful update (DDE links and OLE links only)
            // 15 8000H 1 = Iconified picture link (BIFF8 OLE links only)
            var flags = ReadUInt16(offset);
            offset += 2;

            if (biffVersion >= 5)
            {
                // 0 for global names, or:
                // BIFF5: One-based index to EXTERNSHEET record containing the sheet name,
                // BIFF8: One-based index to sheet list in preceding EXTERNALBOOK record.
                _externalSheetIndex = ReadUInt16(offset);
                offset += 2;

                // Not used.
                _ = ReadUInt16(offset);
                offset += 2;
            }

            // Name (byte string, 8-bit string length, ➜2.5.2). See DEFINEDNAME record (➜5.33) for 
            // list of built-in names, if the built-in flag is set in the option flags above.
            if (biffVersion <= 5)
            {
                _name = ReadByteString(offset, out int nameBytesRead);
                offset += nameBytesRead;
            }
            else
            {
                _name = ReadShortXLUnicodeString(offset, out int nameBytesRead);
                offset += nameBytesRead;
            }

            // For external names and analysis add-in functions.
            // Formula data (RPN token array, ➜3)
            // OR for DDE links and OLE link
            //  Last received results of the DDE or OLE link (constant value array, ➜2.5.8)
            if (ContentOffset + offset < Size)
            {
                // Size of the following formula data (sz)
                var cce = ReadUInt16(offset);
                offset += 2;

                // Formula data (RPN token array)
                _dataOffset = offset;
                _dataLength = cce;
            }
            else
            {
                _dataOffset = offset;
                _dataLength = 0;
            }
        }
        else
        {
            // A - fBuiltIn (1 bit): A bit that specifies whether this record specifies a
            // user-defined or built-in external defined name. The value MUST be 0 if this
            // record specifies a DDE data item, an OLE data item or a UDF reference on a
            // XLL or COM add-in. Otherwise, MUST be one of the following:
            // 0 The external defined name is user-defined.
            // 1 The external defined name is built-in.
            // B - fWantAdvise (1 bit): A bit that specifies whether this record is an
            // automatic DDE data item or OLE data item. MUST be one of the following:
            // 0 The record is an external defined name, a manual DDE data item, a manual
            // OLE data item or a UDF reference on a XLL or COM add-in.
            // 1 The record is either an automatic DDE data item or an automatic OLE data item.
            // C - fWantPict (1 bit): A bit that specifies whether this record's linked data uses
            // a picture format. The value MUST be 0 if this record specifies an external defined name or a UDF reference on a XLL or COM add-in.
            // D - fOle (1 bit): A bit that, together with the value of fOleLink, specifies the
            // structure of body. The value MUST be 0 if this record is an external defined name,
            // an OLE data item or a UDF reference on a XLL or COM add-in. If this value is 1,
            // fOleLink MUST be 0.
            // E - fOleLink (1 bit):  A bit that, together with the value of fOle, specifies
            // the structure of body. The value MUST be 0 if this record is an external
            // defined name or a UDF reference on a XLL or COM add-in. If this value is 1,
            // fOle MUST be 0 and this record specifies an OLE data item.
            // cf (10 bits): A signed integer that specifies the type of the cached
            // clipboard format for a DDE data item or an OLE data item. The value MUST be
            // 0 if this record is an external defined name or a UDF reference on a XLL or
            // COM add-in. The value MUST be one of the values in the following table:
            // -1 There is no cached clipboard format.
            // 0 This record is an external defined name or the cached clipboard format is
            // text. For the text format, each line ends with a carriage return/linefeed
            // (CR-LF) combination. A null character signals the end of the data.
            // 2 Cached clipboard format is Enhanced Metafile.
            // 5 Cached clipboard format is CSV (comma-delimited).
            // 6 Cached clipboard format is Microsoft Symbolic Link (SYLK). SYLK is a
            // format used to exchange data between applications.
            // 7 Cached clipboard format is rich text (RTF).
            // 8 Cached clipboard format is BIFF8.
            // 9 Cached clipboard format is Bitmap.
            // 16 Cached clipboard format is a table created using a specific application<82>.
            // 20 Cached clipboard format is BIFF3.
            // 30 Cached clipboard format is BIFF4.
            // 36 Cached clipboard format is Metafile Picture Format.
            // 44 Cached clipboard format is Unicode text. Each line ends with a carriage
            // return/linefeed (CR-LF) combination. A null character signals the end of
            // the data.
            // 63 Cached clipboard format is BIFF12.
            // F - fIcon (1 bit): A bit that specifies whether linked data is displayed
            // as an icon. The value MUST be 0 if this record is an external defined name,
            // a DDE data item or a UDF reference on a XLL or COM add-in.
            _ = ReadUInt16(offset);
            offset += 2;

            // body (variable):  A variable type field whose type and meaning is dictated
            // by the values of fOle and fOleLink, as specified in the following table:
        }
    }

    public int? ExternalSheetIndex => _externalSheetIndex;

    public string Name => _name;
}
