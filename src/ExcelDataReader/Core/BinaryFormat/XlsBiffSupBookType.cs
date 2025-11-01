namespace ExcelDataReader.Core.BinaryFormat;

/// <summary>
/// Specifies the type of a BIFF8+ external workbook.
/// </summary> 
internal enum XlsBiffSupBookType
{
    Internal3DReference,
    ExternalWorkbookReference,
    AddInFunctionOrDdeOleLink
}
