using Excel.Library;
using Excel.UnitTest.Models;

namespace Excel.UnitTest;

[TestClass]
public class ExcelReadTest
{
    [TestMethod]
    public void ReadSchoolsExcel()
    {
        ExcelLib excelLib = new ExcelLib("./Schools.xlsx");
    }
    [TestMethod]
    public void ReadStudentsExcel()
    {
        ExcelLib excelLib = new ExcelLib("./Students.xlsx");
    }
    [TestMethod]
    public void ReadPatientExcel()
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
                PatientAddress = "789 Pine St",
                PatientAge = 17,
                PatientNumber = "555-9012",
                PatientId = "ID003"
            }
        };

        // Assuming ExcelLib is mocked or reads a predefined Excel file
        ExcelLib excelLib = new ExcelLib("Patients.xlsx");
        List<Patient> actualPatients = excelLib.ReadDataFrame<Patient>();

        // Assertions
        Assert.AreEqual(expectedPatients.Count, actualPatients.Count, "The number of patients should match.");

        for (int i = 0; i < expectedPatients.Count; i++)
        {
            Assert.AreEqual(expectedPatients[i].PatientName, actualPatients[i].PatientName, "Patient names should match.");
            Assert.AreEqual(expectedPatients[i].PatientAddress, actualPatients[i].PatientAddress, "Patient addresses should match.");
            Assert.AreEqual(expectedPatients[i].PatientAge, actualPatients[i].PatientAge, "Patient ages should match.");
            Assert.AreEqual(expectedPatients[i].PatientNumber, actualPatients[i].PatientNumber, "Patient phone numbers should match.");
            Assert.AreEqual(expectedPatients[i].PatientId, actualPatients[i].PatientId, "Patient IDs should match.");
            // Add more assertions as necessary for each property
        }
    }

}