namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Specifies an external sheet.
/// </summary>
internal sealed class XlsBiffExternalSheet : XlsBiffRecord
{
    private readonly string _name;
    private readonly bool _isSelf;
    private readonly List<XlsBiffXti> _refs;

    internal XlsBiffExternalSheet(byte[] bytes, int biffVersion)
        : base(bytes)
    {
        int offset = 0;

        if (biffVersion <= 5)
        {
            // BIFF2-BIFF5 stores ExternSheet differently.
            // OpenOffice specification 5.41 EXTERNSHEET
            // In the file format versions up to BIFF5 this record stores the name of an
            // external document and a sheet name inside of this document. See ➜4.10.1
            // for details about external references in BIFF2-BIFF4 and ➜4.10.2 for BIFF5.
            _name = XlsBiffEncodedUrl.Decode(this, biffVersion, offset, out int bytesRead, out var isSelf);
            offset += bytesRead;

            _isSelf = isSelf;
        }
        else
        {
            // [MS-XLS] 2.5.213 EXTERNSHEET
            // The ExternSheet record specifies a collection of XTI structures.
            // cXTI (2 bytes):  An unsigned integer that specifies the number of elements in the rgXTI array.
            var count = ReadUInt16(offset);
            offset += 2;

            // rgXTI (variable):  An array of XTI structures. The number of elements in this array MUST be cXTI.
            var refs = new List<XlsBiffXti>(count);
            for (int i = 0; i < count; i++)
            {
                refs.Add(new XlsBiffXti(Bytes, ContentOffset + offset));
                offset += 6;
            }

            _refs = refs;

            _isSelf = false;
        }
    }

    public string Name => _name;

    public List<XlsBiffXti> Refs => _refs;

    public bool IsSelf => _isSelf;
}
