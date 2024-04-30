using Excel.Library.Attributes;
using Excel.Library.Helpers;
using Excel.Library.Iterators;
using Excel.Library.Models;
using OfficeOpenXml;
using System.ComponentModel;
using System.Reflection;

namespace Excel.Library;

public partial class ExcelLib
{
    public void WriteDataFrame<T>(List<T> data, string sheetName = "Sheet1",int firstRow = 1, int firstColumn = 1, bool replaceCurrentSheet = true) where T : class
    {
        if (replaceCurrentSheet && SheetExists(_excelPackage,sheetName))
        {
            _excelPackage.Workbook.Worksheets.Delete(sheetName);
        }

        EnsureSheetIsCreated(_excelPackage, sheetName);
        
        List<ExcelProperty> headers = ExcelAttribute.GetExcelWritingProperties<T>();
        ExcelWorksheet sheet = _excelPackage.Workbook.Worksheets[sheetName];
        SheetInfo sheetInfo = new SheetInfo(firstRow,firstColumn,sheet);
        WriteHeaders(headers, sheetInfo);
        WriteBody(data,headers,sheetInfo);
    }
    private void WriteHeaders(List<ExcelProperty> excelAttributes, SheetInfo sheetInfo)
    {
        SheetIterator sheetIterator = new SheetIterator(sheetInfo);
        foreach (var excelAttribute in excelAttributes)
        {
            sheetIterator.GetCurrentCell().Value = excelAttribute.Name;
            sheetIterator.NextColumn();
        }
    }
    private void WriteBody<T>(List<T> data, List<ExcelProperty> headers, SheetInfo sheetInfo) where T : class
    {
        SheetIterator sheetIterator = new SheetIterator(sheetInfo);
        sheetIterator.NextRow();
        foreach (var row in data)
        {
            foreach (var header in headers)
            {
                // Use reflection to get the property based on ExcelProperty's PropertyInfo
                var prop = typeof(T).GetProperty(header.Property.Name);
                if (prop != null)
                {
                    object? value = prop.GetValue(row, null);
                    WriteCell(sheetIterator, prop, value);
                    sheetIterator.NextColumn();
                }

            }
            sheetIterator.ResetColumn();
            sheetIterator.NextRow();
        }
    }
    private void WriteCell(SheetIterator iterator, PropertyInfo propertyInfo,object? value)
    {
        ExcelAttribute? excelAttributes = propertyInfo.GetCustomAttribute<ExcelAttribute>();

        if (excelAttributes == null)
        {
            iterator.GetCurrentCell().Value = value;
            return;
        }
        if(value == null && excelAttributes.DefaultValue != null)
        {
            iterator.GetCurrentCell().Value = excelAttributes.DefaultValue;
            return;
        }
        if (excelAttributes.Type != null && value != null)
        {
            Type targetType = excelAttributes.Type;
            try
            {
                // Attempt to convert value to the specified type.
                TypeConverter converter = TypeDescriptor.GetConverter(targetType);
                if (converter.CanConvertFrom(value.GetType()))
                {
                    var convertedValue = converter.ConvertFrom(value);
                    iterator.GetCurrentCell().Value = convertedValue;
                    return;
                }
            }
            catch
            {
            }
        }

        if (value is string stringValue)
        {
            stringValue = StringHelper.ConvertToCaseStyle(stringValue, excelAttributes.CaseStyle);
            iterator.GetCurrentCell().Value = stringValue;
            return;
        }



        iterator.GetCurrentCell().Value = value?.ToString() ?? value;
    }
    
}
