using Excel.Library.Attributes;

namespace Excel.UnitTest.Models;

public class Student
{
    [Excel(Name = "Student name")]
    public string Name { get; set; } = default!;
    public int Age { get; set; }
    [Excel(DefaultValue = "No address")]
    public string Address { get; set; }
    [Excel(IsProperty = false)]
    public bool IsAbove18 { get => Age > 18; }
}
