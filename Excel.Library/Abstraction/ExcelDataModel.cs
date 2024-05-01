using Excel.Library.Attributes;
using System.ComponentModel;
using System.Reflection;

namespace Excel.Library.Abstraction;

public abstract class ExcelDataModel
{
    public virtual bool IsValid()
    {
        Type type = GetType();
        var properties = type.GetProperties();

        foreach (var property in properties)
        {
            var excelAttributes = property.GetCustomAttribute<ExcelAttribute>();
            // If there's no ExcelAttribute, continue to the next property
            if (excelAttributes == null || excelAttributes.IsProperty != true)
                continue;

            
            if (excelAttributes.Type != null)
            {

                var value = property.GetValue(this);
                if(value == null && !excelAttributes.CanBeNull) return false;

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
