#nullable enable

namespace ExcelDataReader.Core;

internal sealed record Note
{
    public Note(int rowIndex, int columnIndex, string? text)
    {
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Text = text;
    }

    public int RowIndex { get; init; }
    
    public int ColumnIndex { get; init; }
    
    public string? Text { get; set; }
}