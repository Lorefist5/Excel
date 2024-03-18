using Excel.Library.Models;
using OfficeOpenXml;
using System.Diagnostics.Metrics;
using System.Reflection;

namespace Excel.Library.Iterators;

public class SheetIterator : IDisposable
{
    private int _firstRow;
    private int _firstColumn;
    private int _currentRow;
    private int _currentColumn;
    //State of whether this class has been disposed off
    private bool _disposed = false;

    private readonly ExcelWorksheet _excelWorksheet;

    public int CurrentColumn { get => _currentColumn; private set => _currentColumn = value; }
    public int CurrentRow { get => _currentRow; private set => _currentRow = value; }

    public SheetIterator(int firstRow, int firstColumn, ExcelWorksheet excelWorksheet)
    {
        _firstRow = firstRow;
        _firstColumn = firstColumn;
        _excelWorksheet = excelWorksheet;
        _currentColumn = firstColumn;
        CurrentRow = firstRow;
    }
    public SheetIterator(SheetInfo sheetInfo)
    {
        _firstRow = sheetInfo.FirstRow;
        _firstColumn = sheetInfo.FirstColumn;
        _excelWorksheet = sheetInfo.WorkSheet;
        _currentColumn = sheetInfo.FirstColumn;
        CurrentRow = sheetInfo.FirstRow;
    }
    public object? this[int row, int column]
    {
        get
        {
            return _excelWorksheet.Cells[row, column].Value;
        }
        set
        {
            _excelWorksheet.Cells[row, column].Value = value;
        }
    }
    public void ForEachHeader(Action<string,int> action)
    {
        SheetIterator sheetIterator = new SheetIterator(_firstRow, _firstColumn, _excelWorksheet);
        try
        {
            while (sheetIterator.GetCurrentValue() != null)
            {
                string? currentHeaderCellValue = sheetIterator.GetCurrentValue()!.ToString();
                action(currentHeaderCellValue, sheetIterator.CurrentColumn);
                sheetIterator.NextColumn();
            }
        }
        finally
        {
            sheetIterator.Dispose();
        }
    }
    public void ForEachRow(Action<List<RowValue?>> action)
    {
        List<string> headers = new List<string>();
        ForEachHeader((header, currentColumn) =>
        {
            headers.Add(header);
        });

        SheetIterator sheetIterator = new SheetIterator(_firstRow + 1, _firstColumn, _excelWorksheet);
        try
        {
            bool allColumnsNull;
            do
            {
                allColumnsNull = true; // Assume all columns are null until proven otherwise
                List<RowValue?> rowData = new List<RowValue?>();

                for (int column = _firstColumn; column < _firstColumn + headers.Count; column++)
                {
                    sheetIterator.CurrentColumn = column; 
                    object? cellValue = sheetIterator.GetCurrentValue();
                    string? currentHeaderCellValue = this[_firstRow, column]!.ToString();
                    rowData.Add(new RowValue() { HeaderValue = currentHeaderCellValue, Value = cellValue});

                    if (cellValue != null)
                    {
                        allColumnsNull = false;
                    }
                }

                if (!allColumnsNull)
                {
                    action(rowData);
                }

                sheetIterator.NextRow(); 
            }
            while (!allColumnsNull);
        }
        finally
        {
            sheetIterator.Dispose();
        }
    }

    public object? GetCurrentValue()
    {
        return _excelWorksheet.Cells[CurrentRow, CurrentColumn]?.Value;
    }
    public string? GetCurrentHeader()
    {
        return _excelWorksheet.Cells[_firstRow, CurrentColumn].Value?.ToString();
    }
    public ExcelRange GetCurrentCell()
    {
        return _excelWorksheet.Cells[CurrentRow, CurrentColumn];
    }
    public SheetIterator NextRow()
    {
        CurrentRow++;
        return this;
    }
    public SheetIterator PreviousRow()
    {
        CurrentRow--;
        return this;
    }
    public SheetIterator NextColumn()
    {
        _currentColumn++;
        return this;
    }
    public SheetIterator PreviousColumn()
    {
        _currentColumn--;
        return this;
    }
    public SheetIterator ResetIndexes()
    {
        _currentColumn = _firstColumn;
        _currentRow = _firstRow;
        return this;
    }
    public SheetIterator ResetColumn()
    {
        _currentColumn = _firstColumn;
        return this;
    }
    public SheetIterator ResetRow()
    {
        _currentRow = _firstRow;
        return this;
    }
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this); // Prevent finalizer from running
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                // Dispose managed resources.
                // Note: ExcelWorksheet is assumed to be managed and should be disposed of by its own library unless explicitly required.
            }
            _disposed = true;
        }
    }

    ~SheetIterator()
    {
        Dispose(false);
    }
}
