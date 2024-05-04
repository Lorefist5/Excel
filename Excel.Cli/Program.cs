using CommandLine;
using Excel.Cli.Generator;
using OfficeOpenXml;

public class Program
{
    public class Options
    {
        [Option('n', "name", HelpText = "Model name.")]
        public string Name { get; set; }

        [Option('p', "path", HelpText = "Path to the Excel file.")]
        public string Path { get; set; }

        [Option('r', "row", Default = 1, HelpText = "Row where the header starts.")]
        public int Row { get; set; }

        [Option('c', "column", Default = 1, HelpText = "Column where the header starts.")]
        public int Column { get; set; }
        [Option('o', "outputPath", HelpText = "The output path of the model")]
        public string OutputPath { get; set; }
        [Option('s', "sheetName", Default = "Sheet1",HelpText = "The name of the sheet")]
        public string SheetName { get; set; }

        [Option('a', "allSheets", Default = false, HelpText = "Generate models for all sheets.")]
        public bool AllSheets { get; set; }
        [Option('f', "Fix", HelpText = "Fix or generate helper methods in Models.")]
        public string FixModelPath { get; set; }
    }

    public static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        Parser.Default.ParseArguments<Options>(args)
            .WithParsed<Options>(o =>
            {
                try
                {
                    var creator = new ExcelModelCreator();


                    if (!string.IsNullOrWhiteSpace(o.FixModelPath))
                    {
                        creator.GenerateProperties(o.FixModelPath);
                        Console.WriteLine("Properties generated successfully.");
                        return;
                    }
                    
                    if (o.AllSheets)
                    {
                        var models = creator.CreateModelsForAllSheets(o.Path, o.Row, o.Column);
                        foreach (var model in models)
                        {
                            File.WriteAllText($@"{o.OutputPath}\{model.Key}.cs", model.Value);
                        }
                    }
                    else
                    {
                        var model = creator.CreateModel(o.Path, o.Name, o.SheetName, o.Row, o.Column);
                        File.WriteAllText($@"{o.OutputPath}\{o.Name}.cs", model);
                    }
                    Console.WriteLine("Model(s) created successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");
                }
            });
    }
}