using Excel.Library.Attributes;
using Excel.Library.Enums;
using Excel.Library.Helpers;
using Excel.Library.Iterators;
using Excel.Library.Models;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Excel.Library;

public partial class ExcelLib
{
    public List<T> ReadDataFrame<T>(string sheetName, int firstRow = 1, int firstColumn = 1) where T : class, new()
    {
        EnsureExcelExists();
        EnsureSheetExists(_excelPackage,sheetName);
        List<ExcelProperty> properties = ExcelAttribute.GetExcelReadingProperties<T>();
        List<T> results = new List<T>();
        var worksheet = _excelPackage.Workbook.Worksheets[sheetName];
        SheetInfo sheetInfo = new SheetInfo(firstRow,firstColumn, IgnoreHeaderCount, IgnoreLastRowCount,worksheet);
        using(SheetIterator sheetIterator = new SheetIterator(sheetInfo))
        {
            sheetIterator.ForEachRow((rowValue) =>
            {
                Dictionary<PropertyInfo, object?> propertyValues = new Dictionary<PropertyInfo, object?>();
                foreach (var value in rowValue)
                {
                    if (value.HeaderValue != null)
                    {
                        
                        PropertyInfo? propertyInfo = GetColumnAsProperty(properties, value);
                        
                        if (propertyInfo != null)
                        {
                            propertyValues.Add(propertyInfo, value.Value);
                        }
                    }

                }
                results.Add(PopulateData<T>(propertyValues));
            });

            return results;
        }
    }
    public List<T> ReadDataFrame<T>(int firstRow = 1, int firstColumn = 1) where T : class, new()
    {
        List<T> largeDataFrame = new();
        var type = typeof(T);

        var excelSheetAttribute = type.GetCustomAttribute<ExcelSheetAttribute>();
        if (excelSheetAttribute == null)
        {
            return ReadDataFrame<T>("Sheet1", firstRow, firstColumn);
        }

        // If ReadMultiple is true, read data from all specified sheets
        if (excelSheetAttribute.ReadMultiple == true)
        {
            if (excelSheetAttribute.ReadingProperties != null)
            {
                foreach (string currentSheetName in excelSheetAttribute.ReadingProperties)
                {
                    if (this.SheetExists(_excelPackage, currentSheetName))
                    {
                        largeDataFrame.AddRange(ReadDataFrame<T>(currentSheetName, firstRow, firstColumn));
                    }
                }
            }
        }
        else
        {
            // If ReadMultiple is false, read data from the first matching sheet
            string sheetName = GetFirstMatchingSheetName<T>(excelSheetAttribute);
            if (!string.IsNullOrEmpty(sheetName))
            {
                largeDataFrame.AddRange(ReadDataFrame<T>(sheetName, firstRow, firstColumn));
            }
        }
        return largeDataFrame;
    }

    private string GetFirstMatchingSheetName<T>(ExcelSheetAttribute excelSheetAttribute) where T : class, new()
    {
        if (excelSheetAttribute.Index.HasValue)
        {
            // Check by index
            int index = excelSheetAttribute.Index.Value;
            if (index < _excelPackage.Workbook.Worksheets.Count)
            {
                var sheetAtIndex = _excelPackage.Workbook.Worksheets[index];
                if (sheetAtIndex != null)
                {
                    return sheetAtIndex.Name;
                }
            }
        }

        if (!string.IsNullOrEmpty(excelSheetAttribute.Name) && this.SheetExists(_excelPackage, excelSheetAttribute.Name))
        {
            // Check by attribute name
            return excelSheetAttribute.Name;
        }

        if (excelSheetAttribute.ReadingProperties != null)
        {
            // Check by reading properties
            foreach (string propName in excelSheetAttribute.ReadingProperties)
            {
                if (this.SheetExists(_excelPackage, propName))
                {
                    return propName;
                }
            }
        }

        return string.Empty; // Return empty string if no matching sheet is found
    }

    private T PopulateData<T>(Dictionary<PropertyInfo, object?> rowValues) where T : class, new()
    {
        T instance = new T();

        foreach (var entry in rowValues)
        {
            PropertyInfo property = entry.Key;
            object? value = entry.Value;

            var excelAttribute = property.GetCustomAttribute<ExcelAttribute>();
            if (value != null && value is string stringValue && excelAttribute != null && excelAttribute.IgnoreCases != null)
            {
                value = StringHelper.RemoveIgnoreCases(stringValue,excelAttribute.IgnoreCases,excelAttribute.CaseSensitive);
                
            }

            if (value == null)
            {
                object? defaultValue = excelAttribute?.DefaultValue;
                property.SetValue(instance, defaultValue);
            }
            else
            {
                try
                {
                    object? convertedValue = Convert.ChangeType(value, property.PropertyType);
                    property.SetValue(instance, convertedValue);
                }
                catch (Exception ex)
                {
                    // Handle exception as needed

                }
            }
        }

        return instance;
    }


    private PropertyInfo? GetColumnAsProperty(List<ExcelProperty> excelProperties, RowValue columnValue)
    {
        string? columnName = columnValue.HeaderValue;
        int columnIndex = columnValue.HeaderIndex;
        if(columnName == null)
        {
            return null;
        }
        ExcelProperty? property = excelProperties.FirstOrDefault(p =>
        {
            var excelAttribute = p.GetExcelAttributes();
            string? attributeName = excelAttribute?.Name;
            List<string>? readingProperties = excelAttribute?.ReadingProperties?.ToList();
            bool caseSensitive = excelAttribute?.CaseSensitive ?? false;
            TrimMode trimMode = excelAttribute != null ? excelAttribute.TrimMode : TrimMode.FrontAndEnd;

            if(excelAttribute != null && excelAttribute.IndexOfHeader == columnIndex)
            {
                return true;
            }
            if(trimMode == TrimMode.FrontAndEnd)
            {
                columnName = columnName.Trim();
            }
            else if (trimMode == TrimMode.Front)
            {
                columnName.TrimStart();
            }
            else if(trimMode == TrimMode.End)
            {
                columnName.TrimEnd();
            }
            else if(trimMode == TrimMode.All)
            {
                columnName = columnName.Replace(" ", "");
            }

            if (excelAttribute != null && excelAttribute.IgnoreHeaderCases != null && excelAttribute.IgnoreHeaderCases.Count() != 0)
            {
                columnName = columnName.RemoveSubstrings(excelAttribute.IgnoreHeaderCases);
            }


            if (caseSensitive)
            {
                bool namePropertyDetected = 
                p.Property.Name.Equals(columnName, StringComparison.Ordinal) || 
                attributeName?.Equals(columnName, StringComparison.Ordinal) == true;

                if(!namePropertyDetected && readingProperties != null)
                {
                    return readingProperties.Contains(columnName);
                }

                return namePropertyDetected;
            }
            else
            {
                bool noneCaseSensitiveNamePropertyDetected = 
                p.Property.Name.Equals(columnName, StringComparison.OrdinalIgnoreCase) ||
                attributeName?.Equals(columnName, StringComparison.OrdinalIgnoreCase) == true;
                if (!noneCaseSensitiveNamePropertyDetected && readingProperties != null)
                {
                    return readingProperties!.Any(s => s.Equals(columnName, StringComparison.OrdinalIgnoreCase));
                }
                return noneCaseSensitiveNamePropertyDetected;
                
            }
        });


        return property?.Property;
    }
}
