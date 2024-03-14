using Excel.Library.Attributes;
using OfficeOpenXml;
using System.Reflection;

namespace Excel.Library;

public partial class ExcelLib
{
    private readonly string excelPath;
    private readonly ExcelPackage _excelPackage;
    public ExcelLib(string excelPath)
    {
        this.excelPath = excelPath;
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        _excelPackage = new ExcelPackage(excelPath);
    }

    public void SaveAs(string newPath)
    {
        _excelPackage.SaveAs(newPath);
    }
    public void Save()
    {
        _excelPackage.Save();
    }



}