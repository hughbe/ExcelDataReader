using System.Text;

namespace ExcelDataReader.Core.BinaryFormat;

internal struct XlsFormulaReaderContext
{
    public List<XlsBiffDefinedName> DefinedNames;
    public List<XlsBiffSupBook> ExternalWorkbooks;
    public List<XlsBiffExternalSheet> ExternalSheets;
    public List<XlsBiffExternalName> ExternalNames;
    public List<XlsBiffArray> Arrays;
    public List<XlsBiffDataTable> DataTables;
    public List<XlsBiffSharedFormula> SharedFormulas;
    public List<XlsBiffBoundSheet> Sheets;
    public Encoding Encoding;

    public XlsFormulaReaderContext()
    {
    }

    public string GetDefindName(uint nameIndex)
    {
        if (DefinedNames == null)
        {
            throw new InvalidOperationException("Defined names collection is required");
        }

        if (nameIndex == 0 || nameIndex > DefinedNames.Count)
        {
            throw new InvalidOperationException("Invalid defined name index.");
        }

        return DefinedNames[(int)nameIndex - 1].Name;
    }

    public XlsBiffExternalSheet GetExternalSheet5(int sheetIndex)
    {
        // In BIFF2-BIFF5, worksheets are stored as a list of EXTERNSHEET records.
        // BIFF2-BIFF5 external sheet indexes are one-based and negative indexes
        // represent internal sheets.
        int actualIndex = Math.Abs(sheetIndex);

        if (ExternalSheets == null)
        {
            throw new InvalidOperationException("External sheets collection is required.");
        }

        if (actualIndex == 0 || actualIndex > ExternalSheets.Count)
        {
            throw new InvalidOperationException("Invalid external sheet index.");
        }

        return ExternalSheets[actualIndex - 1];
    }

    public (XlsBiffSupBook Workbook, XlsBiffXti ExternalReference) GetExternalSheet8(int sheetIndex)
    {
        // In BIFF8, worksheets are stored as a single EXTERNSHEET record.
        // This record contains a list of XTI structures.
        // Sheet references are zero-based.
        if (ExternalSheets == null)
        {
            throw new InvalidOperationException("External sheets collection is required.");
        }

        if (ExternalSheets.Count != 1)
        {
            throw new InvalidOperationException("Invalid external sheets collection. Expected a single EXTERNSHEET record.");
        }

        XlsBiffExternalSheet externSheet = ExternalSheets[0];
        if (sheetIndex < 0 || sheetIndex >= externSheet.Refs.Count)
        {
            throw new InvalidOperationException("Invalid external sheet index.");
        }

        XlsBiffXti xti = externSheet.Refs[sheetIndex];
        XlsBiffSupBook workbook = GetExternalWorkbook(xti.WorkbookIndex);
        return (workbook, xti);
    }

    public XlsBiffSupBook GetExternalWorkbook(int supBookIndex)
    {
        if (ExternalWorkbooks == null)
        {
            throw new InvalidOperationException("External workbooks collection is required.");
        }

        if (supBookIndex < 0 || supBookIndex >= ExternalWorkbooks.Count)
        {
            throw new InvalidOperationException("Invalid external workbook index.");
        }

        // External workbooks are zero-based.
        return ExternalWorkbooks[supBookIndex];
    }

    public string GetExternalName(uint nameIndex)
    {
        if (ExternalNames == null)
        {
            throw new InvalidOperationException("External names collection is required");
        }

        if (nameIndex == 0 || nameIndex > ExternalNames.Count)
        {
            throw new InvalidOperationException("Invalid external name index.");
        }

        // External names are one-based.
        return ExternalNames[(int)nameIndex - 1].Name;
    }
}
