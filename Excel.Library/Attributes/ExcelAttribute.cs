using Excel.Library.Enums;
using Excel.Library.Models;
using System.Reflection;

namespace Excel.Library.Attributes;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
public class ExcelAttribute : Attribute
{
    public string? Name { get; set; } // Column name that will be read for this property and when writting the Dataframe
    public bool IsReadProperty { get; set; } = true; // This property Will be read as a column
    public bool IsWriteProperty { get; set; } = true; // This property will written as a column
    public int IndexOrder { get; set; } = 100; // The order which the column will be written
    public bool IsProperty
    {
        get => IsReadProperty && IsWriteProperty;
        set
        {
            IsReadProperty = value;
            IsWriteProperty = value;
        }
    }
    public object? DefaultValue { get; set; } // If the value that is being written or being read is null it will replace it with this
    public bool CaseSensitive { get; set; } = true; // If the CaseSensitive is off it will read the columns without caring for the case
    public string[]? IgnoreCases { get; set; } //Ignores some cases, if you put "Los" or "The" it will delete themm ex. Los Santos will be = to Santos
    public string[]? ReadingProperties { get; set; } // When searching for the columns in the excel it will also check for all these properties
    public CaseStyle CaseStyle { get; set; } = CaseStyle.Default; //In what case style it will be written
    public static List<ExcelProperty> GetExcelReadingProperties<T>() where T : class
    {
        List<PropertyInfo> properties = typeof(T).GetProperties().ToList();
        List<ExcelProperty> excelProperties = properties.Select(p => new ExcelProperty() { Property = p }).Where(p => p.IsReadProperty).ToList();
        return excelProperties;
    }
    public static List<ExcelProperty> GetExcelWritingProperties<T>() where T : class
    {
        List<PropertyInfo> properties = typeof(T).GetProperties().ToList();
        List<ExcelProperty> excelProperties = properties.Select(p => new ExcelProperty() { Property = p }).Where(p => p.IsWriteProperty).ToList();
        excelProperties = excelProperties.OrderBy(p => p.GetExcelAttributes()?.IndexOrder).ToList();
        return excelProperties;
    }
}