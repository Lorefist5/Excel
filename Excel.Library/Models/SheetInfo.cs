using OfficeOpenXml;

namespace Excel.Library.Models;

public class SheetInfo
{
    public int FirstRow;
    public int FirstColumn;
    public ExcelWorksheet WorkSheet { get; }
    public SheetInfo(int firstRow, int firstColumn, ExcelWorksheet sheet)
    {
        FirstRow = firstRow;
        FirstColumn = firstColumn;
        WorkSheet = sheet;
    }
}
