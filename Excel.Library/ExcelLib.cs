using Excel.Library.Attributes;
using OfficeOpenXml;
using System.Reflection;

namespace Excel.Library;

public partial class ExcelLib
{
    private string _excelPath;
    private ExcelPackage _excelPackage;
    public ExcelLib(string excelPath)
    {
        this._excelPath = EnsureCorrectExtension(excelPath);
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        _excelPackage = new ExcelPackage(excelPath);
    }
    public void ChangeExcel(string path)
    {
        _excelPackage = new ExcelPackage(path);
        _excelPath = EnsureCorrectExtension(path);
    }
    public void SaveAs(string newPath)
    {
        _excelPath = EnsureCorrectExtension(newPath);
        _excelPackage.SaveAs(_excelPath);
    }
    public void Save()
    {
        _excelPackage.Save();
    }
    public string EnsureCorrectExtension(string path)
    {
        string newPath = path;

        if (!Path.HasExtension(newPath))
        {
            newPath += ".xlsx";

        }
        else if (Path.GetExtension(newPath) != "xlsx")
        {
            newPath = Path.ChangeExtension(newPath, "xlsx");
        }
        return newPath;
    }

    private bool ExcelExist()
    {
        return File.Exists(_excelPath);
    }
    private void EnsureExcelExists()
    {
        if (!ExcelExist())
        {
            throw new FileNotFoundException("The excel sheet does not exist");
        }
    }
    private bool SheetExists(ExcelPackage excelPackage, string sheetName)
    {
        return excelPackage.Workbook.Worksheets.Any(x => x.Name == sheetName);
    }
    private void EnsureSheetExists(ExcelPackage excelPackage, string sheetName)
    {
        if (!SheetExists(excelPackage, sheetName))
        {
            throw new Exception("The sheet does not exist");
        }
    }
    private void EnsureSheetIsCreated(ExcelPackage excelPackage, string sheetName)
    {
        if(!SheetExists(excelPackage, sheetName))
        {
            excelPackage.Workbook.Worksheets.Add(sheetName);
        }
    }
}