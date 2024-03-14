using Excel.Library.Attributes;
using Excel.Library.Iterators;
using Excel.Library.Models;
using System.Reflection;

namespace Excel.Library;

public partial class ExcelLib
{
    public List<T> ReadDataFrame<T>(string sheetName = "Sheet1", int firstRow = 1, int firstColumn = 1) where T : class, new()
    {
        List<PropertyInfo> properties = typeof(T).GetProperties().ToList();
        List<T> results = new List<T>();
        var worksheet = _excelPackage.Workbook.Worksheets[sheetName];

        SheetIterator sheetIterator = new SheetIterator(firstRow, firstColumn, worksheet);


        sheetIterator.ForEachRow((rowValue) =>
        {
            Dictionary<PropertyInfo, object?> propertyValues = new Dictionary<PropertyInfo, object?>();
            foreach(var value in rowValue)
            {
                if(value.HeaderValue != null)
                {
                    PropertyInfo? propertyInfo = GetColumnAsProperty(properties, value.HeaderValue);
                    if(propertyInfo != null)
                    {
                        propertyValues.Add(propertyInfo, value.Value);
                        
                    }
                }
                
            }
            results.Add(PopulateData<T>(propertyValues));
        });

        return results;
    }

    private T PopulateData<T>(Dictionary<PropertyInfo, object?> rowValues) where T : class, new()
    {
        T instance = new T(); 

        foreach (var entry in rowValues)
        {
            PropertyInfo property = entry.Key;
            object? value = entry.Value;

            if(value == null)
            {
                object? defaultValue = property.GetCustomAttribute<ExcelAttribute>()?.DefaultValue;
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
                }
            }

        }

        return instance;
    }

    private PropertyInfo? GetColumnAsProperty(List<PropertyInfo> properties, string columnName)
    {
        
        var excelProperties = properties.Where(p => p.GetCustomAttribute<ExcelAttribute>()?.IsProperty != false).ToList();
        PropertyInfo? property = excelProperties.FirstOrDefault(p => p.Name == columnName || p.GetCustomAttribute<ExcelAttribute>()?.Name == columnName);

        return property;
    }
}
