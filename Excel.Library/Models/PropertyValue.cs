using Excel.Library.Attributes;
using System.Reflection;

namespace Excel.Library.Models;

public class PropertyValue
{
    public ExcelProperty Property { get; set; } = default!;
    public object? Value { get; set; }  

}
