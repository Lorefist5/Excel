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
    [TestMethod]
    public void TestExcelSheetName()
    {
        // Arrange
        var inputFolder = Enviroment.InputFolderPath;
        var excelLib = new ExcelLib(Path.Combine(inputFolder, "SheetNameTest"));

        // Act
        var dataFrames = excelLib.ReadDataFrame<Patient>();

        // Define expected results
        var expectedPatientsSheet1 = new Patient
        {
            PatientId = "ID001",
            PatientName = "From sheet 1",
            PatientAddress = "123 Main St",
            PatientAge = "25",
            PatientNumber = "555-1234",
        };

        var expectedPatientsPatients = new Patient
        {
            PatientId = "ID001",
            PatientName = "From Patients",
            PatientAddress = "123 Main St",
            PatientAge = "25",
            PatientNumber = "555-1234",
        };

        var expectedPatientsMySheetName = new Patient
        {
            PatientId = "ID001",
            PatientName = "From mySheetName",
            PatientAddress = "123 Main St",
            PatientAge = "25",
            PatientNumber = "555-1234",
        };

        var expectedPatientsTestSheet = new Patient
        {
            PatientId = "ID001",
            PatientName = "From test",
            PatientAddress = "123 Main St",
            PatientAge = "25",
            PatientNumber = "555-1234",
        };

        // Act
        var firstDataFrame = excelLib.ReadDataFrame<Patient>("Sheet1");
        var secondDataFrame = excelLib.ReadDataFrame<Patient>("Patients");
        var thirdDataFrame = excelLib.ReadDataFrame<Patient>("MySheetName");
        var fourthDataFrame = excelLib.ReadDataFrame<Patient>("TestSheet");

        // Assert
        Assert.AreEqual(firstDataFrame.Count, 1, "The number of patients in Sheet1 should be 1.");
        Assert.AreEqual(secondDataFrame.Count, 1, "The number of patients in Patients should be 1.");
        Assert.AreEqual(thirdDataFrame.Count, 1, "The number of patients in MySheetName should be 1.");
        Assert.AreEqual(fourthDataFrame.Count, 1, "The number of patients in TestSheet should be 1.");

        firstDataFrame.Should().ContainEquivalentOf(expectedPatientsSheet1, because: "all patient details from Sheet1 should match.");
        secondDataFrame.Should().ContainEquivalentOf(expectedPatientsPatients, because: "all patient details from Patients should match.");
        thirdDataFrame.Should().ContainEquivalentOf(expectedPatientsMySheetName, because: "all patient details from MySheetName should match.");
        fourthDataFrame.Should().ContainEquivalentOf(expectedPatientsTestSheet, because: "all patient details from TestSheet should match.");
    }

}