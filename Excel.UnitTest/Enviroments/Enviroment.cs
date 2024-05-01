namespace Excel.UnitTest.Enviroments;

public static class Enviroment
{
    public static string ProjectFolderPath => Path.Combine(Directory.GetCurrentDirectory(), @"..\..\..\");
    public static string InputFolderPath => Path.Combine(ProjectFolderPath, "Inputs");
}
