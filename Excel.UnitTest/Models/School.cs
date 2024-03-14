using Excel.Library.Attributes;

namespace Excel.UnitTest.Models;

public class School
{
    [Excel(Name = "School Name")]
    public string Name { get; set; }
    public string SchoolDirector { get; set; }
    
}
