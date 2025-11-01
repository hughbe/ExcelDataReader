using System.Globalization;
using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Specifies the number of external sheets.
/// </summary>
internal sealed class XlsBiffExternalSheetCount : XlsBiffRecord
{
    internal XlsBiffExternalSheetCount(byte[] bytes)
        : base(bytes)
    {
        // OpenOffice specification 5.40 EXTERNCOUNT
        // This record contains the number of following EXTERNSHEET records. In BIFF8
        // this record is omitted because there occurs only one EXTERNSHEET record.
        // See ➜4.10.1 for details about external references in BIFF2-BIFF4 and
        // ➜4.10.2 for BIFF5.
        // Number of following EXTERNSHEET records.
        Count = ReadUInt16(0);
    }

    public ushort Count { get; }
}
