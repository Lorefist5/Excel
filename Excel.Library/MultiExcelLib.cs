using Excel.Library.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Library;

public class MultiExcelLib
{
    private List<ExcelLibInformation> _excelLibs;
    public MultiExcelLib(string filePath)
    {
        _excelLibs = new List<ExcelLibInformation> { new ExcelLibInformation { ExcelPath = filePath } };
        
    }
    public MultiExcelLib(ExcelLibInformation excelLib)
    {
        _excelLibs = new List<ExcelLibInformation> { excelLib };
    }
    public MultiExcelLib(List<string> excels)
    {

       _excelLibs = excels.Select(x => new ExcelLibInformation { ExcelPath = x }).ToList();
    }
    public MultiExcelLib(List<ExcelLibInformation> excelLibs)
    {
        _excelLibs = excelLibs;
    }


    public void AddExcelLib(ExcelLibInformation excelLib)
    {
        _excelLibs.Add(excelLib);
    }
    public void AddExcelLib(string filePath)
    {
        _excelLibs.Add(new ExcelLibInformation { ExcelPath = filePath });
    }
    public void AddExcelLib(ExcelLib excelLib)
    {

       _excelLibs.Add(new ExcelLibInformation { ExcelPath = excelLib.ExcelPath, IgnoreHeaderCount = excelLib.IgnoreHeaderCount, IgnoreLastRowCount = excelLib.IgnoreLastRowCount });
    }
    public void AddExcelLibs(List<ExcelLibInformation> excelLibs)
    {
        _excelLibs.AddRange(excelLibs);
    }
    public void AddExcelLibs(List<string> filePaths)
    {
        _excelLibs.AddRange(filePaths.Select(x => new ExcelLibInformation { ExcelPath = x }));
    }
    public void AddExcelLibs(List<ExcelLib> excelLibs)
    {
        _excelLibs.AddRange(excelLibs.Select(x => new ExcelLibInformation { ExcelPath = x.ExcelPath, IgnoreHeaderCount = x.IgnoreHeaderCount, IgnoreLastRowCount = x.IgnoreLastRowCount }));
    }
    public void AddExcelLibs(List<string> filePaths, int ignoreHeaderCount, int ignoreLastRowCount)
    {
        _excelLibs.AddRange(filePaths.Select(x => new ExcelLibInformation { ExcelPath = x, IgnoreHeaderCount = ignoreHeaderCount, IgnoreLastRowCount = ignoreLastRowCount }));
    }

    public void WriteDataFrame<T>(List<T> data, string sheetName = "Sheet1", int firstRow = 1, int firstColumn = 1, bool replaceCurrentSheet = true) where T : class
    {
        foreach (var excelLib in _excelLibs)
        {
            ExcelLib lib = new(excelLib.ExcelPath);
            lib.IgnoreHeaderCount = excelLib.IgnoreHeaderCount;
            lib.IgnoreLastRowCount = excelLib.IgnoreLastRowCount;
            lib.WriteDataFrame(data, sheetName, firstRow, firstColumn, replaceCurrentSheet);
            lib.ChangeExcel(excelLib.OutputPath);
            lib.Save();
        }
    }
    public IEnumerable<T> ReadDataFrame<T>() where T : class, new()
    {
        List<T> results = new();
        foreach (var excelLib in _excelLibs)
        {
            ExcelLib lib = new(excelLib.ExcelPath);
            lib.IgnoreHeaderCount = excelLib.IgnoreHeaderCount;
            lib.IgnoreLastRowCount = excelLib.IgnoreLastRowCount;
            results.AddRange(lib.ReadDataFrame<T>(excelLib.FirstRow, excelLib.FirstColumn));
        }
        return results;
    }
    public IEnumerable<T> ReadDataFrame<T>(string sheetName) where T : class, new()
    {
        List<T> results = new();
        foreach (var excelLib in _excelLibs)
        {
            ExcelLib lib = new(excelLib.ExcelPath);
            lib.IgnoreHeaderCount = excelLib.IgnoreHeaderCount;
            lib.IgnoreLastRowCount = excelLib.IgnoreLastRowCount;
            results.AddRange(lib.ReadDataFrame<T>(sheetName, excelLib.FirstRow, excelLib.FirstColumn));
        }
        return results;
    }
}
