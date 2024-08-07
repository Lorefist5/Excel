using Excel.Library.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.UnitTest.Models;

[ExcelSheet(Name = "GenericExcelSchoolReport", ReadingProperties = ["Informe de Vacunación 2023-24"])]
public class GenericExcelSchoolReport : ReportBase
{
    [Excel(Name = "School Type", ReadingProperties = ["School Type", "Grado Auditado"], IgnoreHeaderCases = [":"], CanBeNull = false)]
    public string SchoolType { get; set; } = default!;

    public override Dictionary<string, int> GetAntigensWithValues()
    {
        throw new NotImplementedException();
    }
    public override bool IsValid()
    {
        return !string.IsNullOrWhiteSpace(SchoolPin);
    }

    public override string GetSchoolType()
    {
        //If it contains K then it is a Kindergarten,
        //If it contains HS then it is a High School,
        //If it contains CCD then it is a DAYCARE (CCD),
        //If it contains 7 then its a 7MO
        //If it contains 8 then its a 8VO
        //If it contains UNI then its a UNIVERSIDAD
        //Based on SchoolType property
        if (SchoolType.Contains("K"))
        {
            return "Kindergarten";
        }
        else if (SchoolType.Contains("HS"))
        {
            return "High School";
        }
        else if (SchoolType.Contains("CCD"))
        {
            return "DAYCARE (CCD)";
        }
        else if (SchoolType.Contains("7"))
        {
            return "7th Grade";
        }
        else if (SchoolType.Contains("8"))
        {
            return "8th Grade";
        }
        else if (SchoolType.Contains("UNI"))
        {
            return "University";
        }
        else
        {
            return "Generic School";
        }

    }
}
