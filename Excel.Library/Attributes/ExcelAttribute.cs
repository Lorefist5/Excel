namespace Excel.Library.Attributes;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
public class ExcelAttribute : Attribute
{
    public ExcelAttribute(string? Name = null, bool? IsProperty = null, object? defaultValue = null)
    {
        this.Name = Name;
        this.IsProperty = IsProperty;
        DefaultValue = defaultValue;
    }

    public string? Name { get; set; }
    public bool? IsProperty { get; set; }
    public object? DefaultValue { get; set; }
}
