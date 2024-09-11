using Excel.Library.Attributes;
using Excel.Library.Enums;
using System.ComponentModel.DataAnnotations.Schema;
using System.Reflection;

namespace Excel.Library.Models;

public class ExcelProperty
{
    public ExcelProperty()
    {
        
    }
    public ExcelProperty(PropertyInfo property)
    {
        Property = property;
    }
    public PropertyInfo Property { get; set; } = default!;
    public string Name { 
        get
        {
            if (GetExcelAttributes() != null && GetExcelAttributes().Name != null)
            {
                return GetExcelAttributes()!.Name!;
            }
            else if (Property.GetCustomAttribute<ColumnAttribute>() != null)
            {
                return Property.GetCustomAttribute<ColumnAttribute>()!.Name!;
            }
            else
                return Property.Name;
        }
    }
    public bool HasExcelProperties
    {
        get => Property.GetCustomAttribute<ExcelAttribute>() != null;
    }
    public bool IsReadProperty
    {
        get => Property.GetCustomAttribute<ExcelAttribute>()?.IsReadProperty != false || !string.IsNullOrEmpty(Property.GetCustomAttribute<ColumnAttribute>()?.Name); // Returns true if its null or IsPropery is true
    }
    public bool IsWriteProperty
    {
        get => Property.GetCustomAttribute<ExcelAttribute>()?.IsWriteProperty != false || !string.IsNullOrEmpty(Property.GetCustomAttribute<ColumnAttribute>()?.Name); // Returns true if its null or IsPropery is true

    }
    
    public ExcelAttribute? GetExcelAttributes()
    {
        return Property.GetCustomAttribute<ExcelAttribute>();
    }
    
    public CaseStyle CaseStyle
    {
        get
        {
            if (!IsReadProperty)return CaseStyle.Default;
            else return Property.GetCustomAttribute<ExcelAttribute>()!.CaseStyle;
        }
    }

}
