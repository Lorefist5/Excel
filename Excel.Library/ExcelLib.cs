﻿using Excel.Library.Attributes;
using OfficeOpenXml;
using System.Reflection;

namespace Excel.Library;

public partial class ExcelLib
{
    public string ExcelPath => _excelPath;
    private string _excelPath;
    private ExcelPackage _excelPackage;
    private int _ignoreHeaderCount;
    private int _ignoreLastRowCount;
    public int IgnoreHeaderCount
    {
        get => _ignoreHeaderCount;
        set => _ignoreHeaderCount = value <= 0 ? 1 : value;
    }
    public int IgnoreLastRowCount
    {
        get => _ignoreLastRowCount;
        set => _ignoreLastRowCount = value <= 0 ? 1 : value;
    }

    public ExcelLib(string excelPath)
    {
        this._excelPath = EnsureCorrectExtension(excelPath);
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        _excelPackage = new ExcelPackage(_excelPath);
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
    public void SaveWithPassword(string password)
    {
        _excelPackage.SaveAs(new FileInfo(_excelPath), password);
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