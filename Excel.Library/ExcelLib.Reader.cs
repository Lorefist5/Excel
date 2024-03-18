using Excel.Library.Attributes;
using Excel.Library.Helpers;
using Excel.Library.Iterators;
using Excel.Library.Models;
using System.Reflection;

namespace Excel.Library;

public partial class ExcelLib
{
    public List<T> ReadDataFrame<T>(string sheetName = "Sheet1", int firstRow = 1, int firstColumn = 1) where T : class, new()
    {
        EnsureExcelExists();
        EnsureSheetExists(_excelPackage,sheetName);
        List<ExcelProperty> properties = ExcelAttribute.GetExcelReadingProperties<T>();
        List<T> results = new List<T>();
        var worksheet = _excelPackage.Workbook.Worksheets[sheetName];
        
        using(SheetIterator sheetIterator = new SheetIterator(firstRow, firstColumn, worksheet))
        {
            sheetIterator.ForEachRow((rowValue) =>
            {
                Dictionary<PropertyInfo, object?> propertyValues = new Dictionary<PropertyInfo, object?>();
                foreach (var value in rowValue)
                {
                    if (value.HeaderValue != null)
                    {
                        PropertyInfo? propertyInfo = GetColumnAsProperty(properties, value.HeaderValue);
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


    private PropertyInfo? GetColumnAsProperty(List<ExcelProperty> excelProperties, string columnName)
    {
        
        ExcelProperty? property = excelProperties.FirstOrDefault(p =>
        {
            var excelAttribute = p.GetExcelAttributes();
            string? attributeName = excelAttribute?.Name;
            List<string>? readingProperties = excelAttribute?.ReadingProperties?.ToList();
            bool caseSensitive = excelAttribute?.CaseSensitive ?? false;

            if (caseSensitive)
            {
                bool namePropertyDecected = 
                p.Property.Name.Equals(columnName, StringComparison.Ordinal) || 
                attributeName?.Equals(columnName, StringComparison.Ordinal) == true;

                if(!namePropertyDecected && readingProperties != null)
                {
                    return readingProperties.Contains(columnName);
                }

                return namePropertyDecected;
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
