using Excel.Library.Attributes;
using System.ComponentModel;
using System.Reflection;

namespace Excel.Library.Abstraction;

public abstract class ExcelDataModel
{
    public virtual bool IsValid()
    {
        Type type = GetType();
        var properties = type.GetProperties().Where(p => p.GetCustomAttribute<ExcelAttribute>() != null && p.GetCustomAttribute<ExcelAttribute>()!.IsProperty != false).ToList();

        foreach (var property in properties)
        {
            var excelAttributes = property.GetCustomAttribute<ExcelAttribute>();


            if (excelAttributes.Type != null)
            {

                var value = property.GetValue(this);
                if (value == null && !excelAttributes.CanBeNull) return false;

                if (value != null && !TryConvert(value, excelAttributes.Type))
                {
                    return false;
                }
            }
        }

        return true;
    }
    private bool TryConvert(object value, Type targetType)
    {
        try
        {

            TypeConverter converter = TypeDescriptor.GetConverter(targetType);
            if (converter != null && converter.CanConvertFrom(value.GetType()))
            {
                var convertedValue = converter.ConvertFrom(value);
                return true;
            }
        }
        catch
        {
        }

        return false;
    }

}
