using Excel.UnitTest.Enviroments;
using System.Security.Cryptography.X509Certificates;

namespace Excel.UnitTest;

[TestClass]
public class ExcelReadTest
{

    [TestMethod]
    public void TestRowSpacing()
    {
        Defaults.IgnoreHeaderCount = 3;
        var inputFolder = Enviroment.InputFolderPath;

        var excelLib = new ExcelLib(Path.Combine(inputFolder, "HeaderCountTest.xlsx"));

        var firstDataFrame = excelLib.ReadDataFrame<Patient>();
        var secondDataFrame = excelLib.ReadDataFrame<Patient>(firstRow: 10);

        Assert.AreEqual(firstDataFrame.Count, secondDataFrame.Count, "The number of students should match.");
        for (int i = 0; i < firstDataFrame.Count; i++)
        {
            Assert.AreEqual(firstDataFrame[i].PatientName, secondDataFrame[i].PatientName, "Patient names should match.");
            Assert.AreEqual(firstDataFrame[i].PatientAddress, secondDataFrame[i].PatientAddress, "Patient addresses should match.");
            Assert.AreEqual(firstDataFrame[i].PatientAge, secondDataFrame[i].PatientAge, "Patient ages should match.");
            Assert.AreEqual(firstDataFrame[i].PatientNumber, secondDataFrame[i].PatientNumber, "Patient phone numbers should match.");
            Assert.AreEqual(firstDataFrame[i].PatientId, secondDataFrame[i].PatientId, "Patient IDs should match.");
        } 
    }
    [TestMethod]
    public void TestColumnSpacing()
    {
        Defaults.IgnoreLastRowCount = 5;
        var inputFolder = Enviroment.InputFolderPath;

        var excelLib = new ExcelLib(Path.Combine(inputFolder, "RowSpacingTest.xlsx"));

        var firstDataFrame = excelLib.ReadDataFrame<Patient>();
        var secondDataFrame = excelLib.ReadDataFrame<Patient>(firstRow: 10);

        Assert.AreEqual(firstDataFrame.Count, secondDataFrame.Count, "The number of students should match.");
        for (int i = 0; i < firstDataFrame.Count; i++)
        {
            Assert.AreEqual(firstDataFrame[i].PatientName, secondDataFrame[i].PatientName, "Patient names should match.");
            Assert.AreEqual(firstDataFrame[i].PatientAddress, secondDataFrame[i].PatientAddress, "Patient addresses should match.");
            Assert.AreEqual(firstDataFrame[i].PatientAge, secondDataFrame[i].PatientAge, "Patient ages should match.");
            Assert.AreEqual(firstDataFrame[i].PatientNumber, secondDataFrame[i].PatientNumber, "Patient phone numbers should match.");
            Assert.AreEqual(firstDataFrame[i].PatientId, secondDataFrame[i].PatientId, "Patient IDs should match.");
        }
    }
}