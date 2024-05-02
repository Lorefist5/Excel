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
    private int _ignoreLastRowCount { get; }
    private int _ignoreHeaderCount { get; }
    //State of whether this class has been disposed off
    private bool _disposed = false;

    private readonly ExcelWorksheet _excelWorksheet;

    public int CurrentColumn { get => _currentColumn; private set => _currentColumn = value; }
    public int CurrentRow { get => _currentRow; private set => _currentRow = value; }
    public SheetIterator(SheetInfo sheetInfo)
    {
        _firstRow = sheetInfo.FirstRow;
        _firstColumn = sheetInfo.FirstColumn;
        _excelWorksheet = sheetInfo.WorkSheet;
        _currentColumn = sheetInfo.FirstColumn;
        _ignoreHeaderCount = sheetInfo.IgnoreHeaderCount;
        _ignoreLastRowCount = sheetInfo.IgnoreLastRowCount;

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
        SheetInfo sheetInfo = new SheetInfo(_firstRow, _firstColumn, _ignoreHeaderCount, _ignoreLastRowCount,_excelWorksheet);
        SheetIterator sheetIterator = new SheetIterator(sheetInfo);
        try
        {
            int nullHeadersCount = 0;
            
            while (nullHeadersCount <= _ignoreHeaderCount)
            {
                var value = sheetIterator.GetCurrentValue();
                if (value == null) nullHeadersCount++;
                else
                {
                    nullHeadersCount = 0;
                    string? currentHeaderCellValue = sheetIterator.GetCurrentValue()!.ToString();
                    action(currentHeaderCellValue, sheetIterator.CurrentColumn);
                    
                }
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
        Dictionary<int,string> headers = new ();
        ForEachHeader((header, currentColumn) =>
        {
            headers.Add(currentColumn,header);
        });
        SheetInfo sheetInfo = new SheetInfo(_firstRow + 1, _firstColumn,_ignoreHeaderCount,_ignoreLastRowCount ,_excelWorksheet);
        SheetIterator sheetIterator = new SheetIterator(sheetInfo);
        try
        {
            int nullRowsCount = 0;
            
            do
            {
                nullRowsCount++;
                List<RowValue?> rowData = new List<RowValue?>();

                foreach(var header in headers)
                {
                    sheetIterator.CurrentColumn = header.Key; 
                    object? cellValue = sheetIterator.GetCurrentValue();
                    string? currentHeaderCellValue = this[_firstRow, header.Key]!.ToString();
                    rowData.Add(new RowValue() { HeaderValue = currentHeaderCellValue, Value = cellValue});

                    if (cellValue != null)
                    {
                        nullRowsCount = 0;
                    }
                }

                if (nullRowsCount == 0) // Means that the row wasn't fully null
                {
                    action(rowData);
                }

                sheetIterator.NextRow(); 
            }
            while (nullRowsCount <= _ignoreLastRowCount);
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
