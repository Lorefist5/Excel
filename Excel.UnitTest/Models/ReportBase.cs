using Excel.Library.Abstraction;
using Excel.Library.Attributes;

namespace Excel.UnitTest.Models;

public abstract class ReportBase : ExcelDataModel
{
    //Fetches the school name
    [Excel(Name = "School Name", ReadingProperties = ["Nombre de la Institución"], CaseSensitive = false, IgnoreHeaderCases = [":"], IgnoreCases = ["CCD"], CanBeNull = false)]
    public string Name { get; set; }
    //Fetches the Pin Number

    [Excel(Name = "Pin Number", ReadingProperties = ["Pin", "PinNumber"])]
    public string SchoolPin { get; set; }
    //Fetches the County AKA region of the school
    [Excel(Name = "School Region", ReadingProperties = ["City Location", "County", "Region"], CanBeNull = true)]
    public string? SchoolRegion { get; set; }
    //Fetches the School Director
    [Excel(Name = "Director", ReadingProperties = ["School Director", "Nombre del director o administrador"], IgnoreHeaderCases = [":"], CaseSensitive = false, CanBeNull = true)]
    public string? SchoolDirector { get; set; }
    //Fetches the Enrollment of the school
    [Excel(Name = "Enrollment", ReadingProperties = ["Enrollment CDC"], CaseSensitive = false, Type = typeof(int))]
    public string Enrollment { get; set; } = default!;
    //Fetches the School Address
    [Excel(Name = "School Address", ReadingProperties = ["Address", "Dirección de la Institución"], CanBeNull = true)]
    public string? SchoolAddress { get; set; }
    //Fetches the Excpeted Sample number
    [Excel(Name = "Phone", ReadingProperties = ["Phone Number", "Teléfono"], CanBeNull = true)]
    public string PhoneNumber { get; set; }
    [Excel(Name = "Expected Sample", CanBeNull = true, Type = typeof(int))]
    public string ExpectedExample { get; set; }
    //Fetches the Actual Enrollment number

    [Excel(Name = "Actually Enrollment", ReadingProperties = ["Actual Enrollment"], Type = typeof(int))]
    public string ActuallyEnrollment { get; set; }
    //Fetches the Total Audited number

    [Excel(Name = "Total Audited", Type = typeof(int))]
    public string TotalAudited { get; set; }
    //Fetches the Withdraw number

    [Excel(Name = "Withdraw", Type = typeof(int))]
    public string WithDraw { get; set; }
    //Fetches the Medical Exemption number

    [Excel(Name = "Medical Exemption", Type = typeof(int))]
    public string MedicalExcemptions { get; set; }
    //Fetches the Religious Exemption number

    [Excel(Name = "Religious Exemption", Type = typeof(int))]
    public string ReligiousExcemptions { get; set; }
    [Excel(Name = "Administrator Email", ReadingProperties = ["Correo electrónico de director o administrador"], CanBeNull = true)]
    public string AdministratorEmail { get; set; }

    // None properties
    [Excel(IsProperty = false)]
    public int Enrollmentint => int.Parse(Enrollment);

    [Excel(IsProperty = false)]
    public int ActuallyEnrollmentint => int.Parse(ActuallyEnrollment);

    [Excel(IsProperty = false)]
    public int TotalAuditedint => int.Parse(TotalAudited);

    [Excel(IsProperty = false)]
    public int WithDrawint => int.Parse(WithDraw);

    [Excel(IsProperty = false)]
    public int MedicalExcemptionsint => int.Parse(MedicalExcemptions);

    [Excel(IsProperty = false)]
    public int ReligiousExcemptionsint => int.Parse(ReligiousExcemptions);
    //Makes sure the school has a name and enrollment number
    public override bool IsValid()
    {

        if (string.IsNullOrWhiteSpace(SchoolPin))
        {
            return false;
        }
        if (ActuallyEnrollment == null || ActuallyEnrollment.ToLower() == "closed" || ActuallyEnrollment.ToLower() == "close")
        {
            return false;
        }
        return true;
    }
    public bool IsActive()
    {

        if (ActuallyEnrollment != null && (ActuallyEnrollment.ToLower().Contains("closed") || ActuallyEnrollment.ToLower().Contains("no"))) return false;

        return true;
    }
    public bool IsAudit()
    {
        if (int.TryParse(TotalAudited, out int totalAuditedInt) && totalAuditedInt > 0)
        {
            return true;
        }

        return false;
    }

    public abstract string GetSchoolType();
    public abstract Dictionary<string, int> GetAntigensWithValues();

}


