namespace Excel.Library.Models;

public class ExcelLibInformation
{
    private string? _outputPath = null;
    public required string ExcelPath { get; set; }
    public int IgnoreHeaderCount { get; set; }
    public int IgnoreLastRowCount { get; set; }
    public int FirstRow { get; set; } = 1;
    public int FirstColumn { get; set; } = 1;
    public string OutputPath
    {
        get
        {
            if (_outputPath == null)
            {
                return ExcelPath;
            }
            return _outputPath;
        }
        set
        {
            _outputPath = value;
        }
    }

}
