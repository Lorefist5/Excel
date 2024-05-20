using Excel.Library.Attributes;
using Excel.Library.Enums;
namespace Excel.UnitTest.Models;

[ExcelSheet(Name = "Patient", ReadingProperties = ["Sheet1", "MySheetName", "TestSheet"], ReadMultiple = true)]
public class Patient
{
    [Excel(IgnoreCases = ["Los", "The"], ReadingProperties = ["Name","Patient name"], IndexOrder = 1, CaseSensitive = false, CaseStyle = CaseStyle.PascalCase)]
    public string PatientName { get; set;}

    [Excel(Name = "Address", DefaultValue = "No address")]
    public string? PatientAddress { get; set;}
    [Excel(Name = "Age", CaseSensitive = false, Type = typeof(int))]
    public string PatientAge { get; set;}
    [Excel(Name = "Phone number")]
    public string PatientNumber { get; set; }
    public string PatientId { get; set; }
    

}
