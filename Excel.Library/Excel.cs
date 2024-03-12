using Excel.Library.Attributes;
using OfficeOpenXml;
using System.Reflection;

namespace Excel.Library;

public class Excel
{
    private readonly string excelPath;

    public Excel(string excelPath)
    {
        this.excelPath = excelPath;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    public void SaveAs()
    {

    }
    public void Save()
    {

    }
    public List<T> ReadDataFrame<T>(int firstRow, int firstColumn) where T : class
    {
        Dictionary<int, PropertyInfo> properties = new Dictionary<int, PropertyInfo>();
        //To get the properties you will read the first row until it == null then you will, each time it will check if any of the properties in the model == the name
        throw new NotImplementedException();
    }
    public void WriteDataFrame<T>(List<T> data) where T : class
    {
        throw new NotImplementedException();
    }
    private T PopulateFromRow<T>(Dictionary<int, PropertyInfo> properties, int currentRow) where T : class
    {
        //
        throw new NotImplementedException();
    }
    private PropertyInfo? GetColumnAsProperty(List<PropertyInfo> properties, string columnName)
    {
        var excelProperties = properties.Where(p => p.GetCustomAttribute<ExcelAttribute>().IsProperty != false);
        PropertyInfo? property = excelProperties.FirstOrDefault(p => p.Name == columnName || p.GetCustomAttribute<ExcelAttribute>().Name == columnName);


        return property;
    }

}