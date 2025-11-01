namespace ExcelDataReader.Core.OpenXmlFormat.Records;

internal sealed class CellRecord(int columnIndex, int xfIndex, object value, CellError? error, string formula, string note) : Record
{
    public int ColumnIndex { get; } = columnIndex;

    public int XfIndex { get; } = xfIndex;

    public object Value { get; } = value;

    public CellError? Error { get; } = error;

    public string Formula { get; } = formula;

    public string Note { get; } = note;
}
