using OfficeOpenXml;

namespace Excel.Library.Models;

public class SheetInfo
{
    public int FirstRow;
    public int FirstColumn;

    public int IgnoreHeaderCount { get; }
    public int IgnoreLastRowCount { get; }
    public ExcelWorksheet WorkSheet { get; }
    public SheetInfo(int firstRow, int firstColumn,int ignoreHeaderCount, int ignoreLastRowCount, ExcelWorksheet sheet)
    {
        FirstRow = firstRow;
        FirstColumn = firstColumn;
        IgnoreHeaderCount = ignoreHeaderCount;
        IgnoreLastRowCount = ignoreLastRowCount;
        WorkSheet = sheet;
    }
}
