namespace Excel.UnitTest;

[TestClass]
public class ExcelReadTest
{
    [TestMethod]
    public void ReadSchoolsExcel()
    {
        ExcelLib excelLib = new ExcelLib("./Schools.xlsx");
        List<School> expectedSchool = new List<School>()
        {
            new School() {Name = "San sebastian", SchoolDirector = "Jose Rivera"},
            new School() {Name = "San Jose", SchoolDirector = "Anthonny"}
        };
        
        var actualSchool = excelLib.ReadDataFrame<School>();

        Assert.AreEqual(expectedSchool.Count, actualSchool.Count, "The number of schools should match.");

        for (int i = 0; i < expectedSchool.Count; i++)
        {
            Assert.AreEqual(expectedSchool[i].Name, actualSchool[i].Name, "Patient names should match.");
            Assert.AreEqual(expectedSchool[i].SchoolDirector, actualSchool[i].SchoolDirector, "Patient addresses should match.");
        }
    }
    [TestMethod]
    public void ReadStudentsExcel()
    {
        ExcelLib excelLib = new ExcelLib("./Students.xlsx");
        List<Student> expectedStudents = new List<Student>()
        {
            new Student()
            {
                Name = "Andres",
                Age = 20,
                Address = "No address"
            },
            new Student()
            {
                Name = "The JohnDoe",
                Age = 25,
                Address = "123 Main St"
            },
            new Student()
            {
                Name = "Alice Johnson",
                Age = 17,
                Address = "789 Pine St"
            }
        };

        var actualStudents = excelLib.ReadDataFrame<Student>(firstRow:5,firstColumn:5);
        Assert.AreEqual(expectedStudents.Count, actualStudents.Count, "The number of students should match.");

        for (int i = 0; i < expectedStudents.Count; i++)
        {
            Assert.AreEqual(expectedStudents[i].Name, actualStudents[i].Name, "Patient names should match.");
            Assert.AreEqual(expectedStudents[i].Address, actualStudents[i].Address, "Patient addresses should match.");
            Assert.AreEqual(expectedStudents[i].Age, actualStudents[i].Age, "Patient ages should match.");
        }
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

        ExcelLib excelLib = new ExcelLib("Patients.xlsx");
        List<Patient> actualPatients = excelLib.ReadDataFrame<Patient>();
        actualPatients.Where(p => p.IsAdult);
        // Assertions
        Assert.AreEqual(expectedPatients.Count, actualPatients.Count, "The number of patients should match.");

        for (int i = 0; i < expectedPatients.Count; i++)
        {
            Assert.AreEqual(expectedPatients[i].PatientName, actualPatients[i].PatientName, "Patient names should match.");
            Assert.AreEqual(expectedPatients[i].PatientAddress, actualPatients[i].PatientAddress, "Patient addresses should match.");
            Assert.AreEqual(expectedPatients[i].PatientAge, actualPatients[i].PatientAge, "Patient ages should match.");
            Assert.AreEqual(expectedPatients[i].PatientNumber, actualPatients[i].PatientNumber, "Patient phone numbers should match.");
            Assert.AreEqual(expectedPatients[i].PatientId, actualPatients[i].PatientId, "Patient IDs should match.");
        }
    }
    [TestMethod]
    public void ReadPatientsReadingPropertiesTest()
    {
        ExcelLib excelLib = new ExcelLib("Patients.xlsx");
        List<Patient> firstPatients = excelLib.ReadDataFrame<Patient>();
        excelLib.ChangeExcel("Patients2.xlsx");
        List<Patient> secondPatients = excelLib.ReadDataFrame<Patient>();
        excelLib.ChangeExcel("Patients3.xlsx");
        List<Patient> thirdPatients = excelLib.ReadDataFrame<Patient>();
        for (int i = 0; i < firstPatients.Count; i++)
        {
            Assert.AreEqual(firstPatients[i].PatientName, secondPatients[i].PatientName, "Patient names should match.");
            Assert.AreEqual(firstPatients[i].PatientAddress, secondPatients[i].PatientAddress, "Patient addresses should match.");
            Assert.AreEqual(firstPatients[i].PatientAge, secondPatients[i].PatientAge, "Patient ages should match.");
            Assert.AreEqual(firstPatients[i].PatientNumber, secondPatients[i].PatientNumber, "Patient phone numbers should match.");
            Assert.AreEqual(firstPatients[i].PatientId, secondPatients[i].PatientId, "Patient IDs should match.");
        }
        for (int i = 0; i < firstPatients.Count; i++)
        {
            Assert.AreEqual(firstPatients[i].PatientName, thirdPatients[i].PatientName, "Patient names should match.");
            Assert.AreEqual(firstPatients[i].PatientAddress, thirdPatients[i].PatientAddress, "Patient addresses should match.");
            Assert.AreEqual(firstPatients[i].PatientAge, thirdPatients[i].PatientAge, "Patient ages should match.");
            Assert.AreEqual(firstPatients[i].PatientNumber, thirdPatients[i].PatientNumber, "Patient phone numbers should match.");
            Assert.AreEqual(firstPatients[i].PatientId, thirdPatients[i].PatientId, "Patient IDs should match.");
        }
    }

}