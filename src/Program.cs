using System.CommandLine;

namespace ExcelCsvMerger;

/// <summary>
/// Merges an FBDI Excel document with a group of CSV files.
/// </summary>
class Program
{
    /// <summary>
    /// Entry-point to application.
    /// </summary>
    static async Task<int> Main(string[] args)
    {
        var excelFilePathArgument = new Argument<string>()
        {
            Name = "excelFilePath",
            Description = "File path to Excel document to merge CSV files into."
        };
        var csvFolderPathArgument = new Argument<string>()
        {
            Name = "csvFolderPath",
            Description = "Path to folder with CSV files to merge in. Folder must contain 'csv' files."
        };
        var outputFilePathOption = new Option<string>(
            name: "--output",
            description: "Specify file path to export merged FBDI document to."
        );

        var rootCommand = new RootCommand("Merges FBDI document and CSV files.")
        {
            excelFilePathArgument,
            csvFolderPathArgument,
            outputFilePathOption
        };
        rootCommand.SetHandler(GenerateExcelDocument, excelFilePathArgument, csvFolderPathArgument, outputFilePathOption);

        return await rootCommand.InvokeAsync(args);
    }

    /// <summary>
    /// Validate inputs and generate Excel file.
    /// </summary>
    /// <param name="excelFilePath">Path to Excel file.</param>
    /// <param name="csvFolderPath">Path to folder containing CSVs.</param>
    /// <param name="outputFilePath">Path to output file. Optional.</param>
    private static void GenerateExcelDocument(string excelFilePath, string csvFolderPath, string? outputFilePath)
    {
        if (!InputValidator.ValidateInput(excelFilePath, csvFolderPath))
        {
            return;
        }
        
        outputFilePath ??= $"{Path.GetDirectoryName(excelFilePath)}\\Output{Path.GetExtension(excelFilePath)}";
        if (!TryCreateOutputFile(excelFilePath, outputFilePath))
        {
            return;
        }
        
        ExcelMerger.MergeFiles(outputFilePath, csvFolderPath);
        Console.WriteLine($"Finished generating Excel file: {outputFilePath}");
    }

    /// <summary>
    /// Attempts to create the output file.
    /// </summary>
    private static bool TryCreateOutputFile(string fileName, string outputName)
    {
        Console.WriteLine($"Creating output file: {outputName}");
        try
        {
            File.Copy(fileName, outputName, true);
            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine($"Failed to create output file due to error: {e.Message}");
            return false;
        }
    }
}