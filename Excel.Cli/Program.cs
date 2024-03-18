

using System;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.IO;
using System.Threading.Tasks;

class Program
{
    static async Task<int> Main(string[] args)
    {
        // Define a new root command with a description.
        var rootCommand = new RootCommand("Writes a given sentence to a specified text file.")
        {
            // Define the first argument to capture the text file name.
            new Argument<string>("fileName", "The name of the text file to write to."),
            
            // Define the second argument to capture the sentence to write.
            new Argument<string>("sentence", "The sentence to write into the text file.")
        };

        // Set the handler for the command.
        rootCommand.Handler = new Command()
        {
            // Combine the file name with the current directory (or specify a path).
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), fileName);

            // Write the sentence to the specified file.
            await File.WriteAllTextAsync(filePath, sentence);

            Console.WriteLine($"The sentence has been written to {filePath}");
        });

        // Parse the incoming args and invoke the handler.
        return await rootCommand.InvokeAsync(args);
    }
}
