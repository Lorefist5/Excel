namespace Excel.Library.Attributes;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
public class ExcelAttribute : Attribute
{


    public string? Name { get; set; }
    public bool IsProperty { get; set; } = true;
    public object? DefaultValue { get; set; }
}
