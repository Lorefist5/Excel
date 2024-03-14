using Excel.Library.Attributes;

namespace Excel.UnitTest.Models;

public class Patient
{

    public string PatientName { get; set;}

    [Excel(Name = "Address", DefaultValue = "No address")]
    public string PatientAddress { get; set;}
    [Excel(Name = "Age")]
    public int PatientAge { get; set;}
    [Excel(Name = "Phone number")]
    public string PatientNumber { get; set; }
    public string PatientId { get; set; }

    [Excel(IsProperty = false)]
    public bool IsUnderAge { get => PatientAge <= 18; }
}
