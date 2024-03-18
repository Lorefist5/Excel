using Excel.Library.Attributes;
using Excel.Library.Enums;
namespace Excel.UnitTest.Models;

public class Patient
{
    [Excel(IgnoreCases = ["Los", "The"], ReadingProperties = ["Name","Patient name"], IndexOrder = 1, CaseSensitive = false, CaseStyle = CaseStyle.PascalCase)]
    public string PatientName { get; set;}

    [Excel(Name = "Address", DefaultValue = "No address")]
    public string? PatientAddress { get; set;}
    [Excel(Name = "Age", CaseSensitive = false)]
    public int PatientAge { get; set;}
    [Excel(Name = "Phone number")]
    public string PatientNumber { get; set; }
    public string PatientId { get; set; }
    
    [Excel(IsProperty = false)]
    public bool IsUnderAge { get => PatientAge <= 18; }
    [Excel(IsReadProperty = false, IsWriteProperty = true)]
    public bool IsAdult { get => PatientAge > 18; }
}
