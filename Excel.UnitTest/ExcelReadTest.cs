using Excel.Library.Models;
using Excel.UnitTest.Enviroments;
using OfficeOpenXml.FormulaParsing.Excel.Functions;


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


    //Write mutli excel reader tests
    [TestMethod]
    public void TestMultiExcelReader()
    {
        //Write the test sheets

        // Arrange
        var inputFolder = Enviroment.InputFolderPath;
        //Can you generate dummy data?
        var patients1 = new List<Patient>
        {
            new Patient
            {
                PatientId = "ID001",
                PatientName = "From sheet 1",
                PatientAddress = "123 Main St",
                PatientAge = "25",
                PatientNumber = "555-1234",
            },
            new Patient
            {
                PatientId = "ID002",
                PatientName = "From sheet 2",
                PatientAddress = "123 Main St",
                PatientAge = "25",
                PatientNumber = "555-1234",
            },
            new Patient
            {
                PatientId = "ID003",
                PatientName = "From sheet 3",
                PatientAddress = "123 Main St",
                PatientAge = "25",
                PatientNumber = "555-1234",
            },

        };
        var patients2 = new List<Patient>
        {
            new Patient
            {
                PatientId = "ID004",
                PatientName = "From sheet 4",
                PatientAddress = "456 Pine St",
                PatientAge = "32",
                PatientNumber = "555-6789",
            },
            new Patient
            {
                PatientId = "ID005",
                PatientName = "From sheet 5",
                PatientAddress = "789 Oak St",
                PatientAge = "28",
                PatientNumber = "555-9876",
            },
            new Patient
            {
                PatientId = "ID006",
                PatientName = "From sheet 6",
                PatientAddress = "101 Maple St",
                PatientAge = "45",
                PatientNumber = "555-6543",
            },
        };

        var patients3 = new List<Patient>
        {
            new Patient
            {
                PatientId = "ID007",
                PatientName = "From sheet 7",
                PatientAddress = "234 Birch St",
                PatientAge = "37",
                PatientNumber = "555-3210",
            },
            new Patient
            {
                PatientId = "ID008",
                PatientName = "From sheet 8",
                PatientAddress = "567 Cedar St",
                PatientAge = "29",
                PatientNumber = "555-4321",
            },
            new Patient
            {
                PatientId = "ID009",
                PatientName = "From sheet 9",
                PatientAddress = "890 Willow St",
                PatientAge = "41",
                PatientNumber = "555-5432",
            },
        };


        var excel1 = Path.Combine(inputFolder, "SheetNameTest1");
        var excel2 = Path.Combine(inputFolder, "SheetNameTest2");
        var excel3 = Path.Combine(inputFolder, "SheetNameTest3");
        var excelFiles = new List<string>()
        {
            excel1,
            excel2,
            excel3
        };



        var multiExcelLib = new MultiExcelLib(excelFiles);
        ExcelLib excelLib1 = new(excel1);
        ExcelLib excelLib2 = new(excel2);
        ExcelLib excelLib3 = new(excel3);
        excelLib1.WriteDataFrame(patients1);
        excelLib2.WriteDataFrame(patients2);
        excelLib3.WriteDataFrame(patients3);

        excelLib1.Save();
        excelLib2.Save();
        excelLib3.Save();
        // Act
        var dataFrames = multiExcelLib.ReadDataFrame<Patient>("Sheet1");

        // Define expected results
        var expectedPatientsSheet1 = new Patient
        {
            PatientId = "ID001",
            PatientName = "From sheet 1",
            PatientAddress = "123 Main St",
            PatientAge = "25",
            PatientNumber = "555-1234",
        };

        // Assert
        Assert.AreEqual(dataFrames.Count(), 3, "The number of patients in Sheet1 should be 3.");
        dataFrames.Should().ContainEquivalentOf(expectedPatientsSheet1, because: "all patient details from Sheet1 should match.");
    }
    [TestMethod]
    public void TestMultiRead()
    {
        List<ExcelLibInformation> informes = new();
        string inputFolderPath = @"C:\Users\17874\Source\Repos\Lorefist5\SchoolAuditSelection\Console\Input\";
        MultiExcelLib informesReader;
        if (string.IsNullOrWhiteSpace(inputFolderPath))
        {
            throw new Exception("Input folder path is required");
        }
        foreach (var file in Directory.GetFiles(Path.Combine(inputFolderPath, "Informes"), "*.xlsx"))
        {
            informes.Add(new ExcelLibInformation { ExcelPath = file, FirstRow = 3, IgnoreHeaderCount = 10, IgnoreLastRowCount = 10 });
        }
        informesReader = new(informes);
        
        var data = informesReader.ReadDataFrame<GenericExcelSchoolReport>();
    }
}