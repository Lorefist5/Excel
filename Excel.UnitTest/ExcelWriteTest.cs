namespace Excel.UnitTest;
[TestClass]
public class ExcelWriteTest
{
    [TestMethod]
    public void WriteSchoolsExcel()
    {
    }
    [TestMethod]
    public void WriteStudentsExcel()
    {
    }
    [TestMethod]
    public void WritePatientExcel()
    {
        // Expected patient list
        List<Patient> expectedPatients = new List<Patient>
        {
            new Patient
            {
                PatientName = "John Doe",
                PatientAddress = "123 Main St",
                PatientAge = 25,
                PatientNumber = "555-1234",
                PatientId = "ID001"
            },
            new Patient
            {
                PatientName = "Jane Smith",
                PatientAddress = "456 Elm St",
                PatientAge = 30,
                PatientNumber = "555-5678",
                PatientId = "ID002"
            },
            new Patient
            {
                PatientName = "Alice Johnson",
                PatientAddress = null,
                PatientAge = 17,
                PatientNumber = "555-9012",
                PatientId = "ID003"
            }
        };


        ExcelLib excelLib = new ExcelLib("PatientsWrite.xlsx");
        excelLib.WriteDataFrame(expectedPatients);
        excelLib.SaveAs("PatientsWrite.xlsx");
    }
}
