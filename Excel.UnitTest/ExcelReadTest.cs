using Excel.UnitTest.Enviroments;


namespace Excel.UnitTest;

[TestClass]
public class ExcelReadTest
{

    [TestMethod]
    public void TestRowSpacing()
    {

        var inputFolder = Enviroment.InputFolderPath;

        var excelLib = new ExcelLib(Path.Combine(inputFolder, "HeaderCountTest"));
        excelLib.IgnoreHeaderCount = 3;
        var firstDataFrame = excelLib.ReadDataFrame<Patient>();
        var secondDataFrame = excelLib.ReadDataFrame<Patient>(firstRow: 10);

        Assert.AreEqual(firstDataFrame.Count, secondDataFrame.Count, "The number of students should match.");
        firstDataFrame.Should().HaveCount(secondDataFrame.Count, because: "the number of patients should match in both data frames.");
        firstDataFrame.Should().BeEquivalentTo(secondDataFrame, options => options.ComparingByMembers<Patient>(), because: "all patient details should match.");
    }
    [TestMethod]
    public void TestColumnSpacing()
    {
        List<Patient> patients = new List<Patient>
        {
            new Patient
            {
                PatientId = "ID001",
                PatientName = "John Doe",
                PatientAddress = "123 Main St",
                PatientAge = "25",
                PatientNumber = "555-1234",
            },
            new Patient
            {
                PatientId = "ID002",
                PatientName = "Jane Smith",
                PatientAddress = "456 Elm St",
                PatientAge = "30",
                PatientNumber = "555-5678",
            },
            new Patient
            {
                PatientId = "ID003",
                PatientName = "Alice Johnson",
                PatientAddress = "No address",
                PatientAge = "17",
                PatientNumber = "555-9012"
            },
            new Patient
            {
                PatientId = "ID003", 
                PatientName = "Alice Johnson",
                PatientAddress = "No address",
                PatientAge = "17",
                PatientNumber = "555-9012"
            }
        };


        var inputFolder = Enviroment.InputFolderPath;

        var excelLib = new ExcelLib(Path.Combine(inputFolder, "RowSpacingTest"));
        excelLib.IgnoreLastRowCount = 5;
        var firstDataFrame = excelLib.ReadDataFrame<Patient>();
        
        Assert.AreEqual(firstDataFrame.Count, patients.Count, "The number of students should match.");
        firstDataFrame.Should().HaveCount(patients.Count, because: "the number of patients should match in both data frames.");
        firstDataFrame.Should().BeEquivalentTo(patients, options => options.ComparingByMembers<Patient>(), because: "all patient details should match.");
    }
}