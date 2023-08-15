namespace ExcelCsvMerger;

/// <summary>
/// For validating command line arguments and options.
/// </summary>
public static class InputValidator
{
    /// <summary>
    /// Validates input arguments.
    /// </summary>
    public static bool ValidateInput(string excelFilePath, string csvFolderPath)
    {
        var excelFilePathValid = ValidateExcelFile(excelFilePath);
        var csvFolderPathValid = ValidateCsvFolder(csvFolderPath);
        
        return excelFilePathValid && csvFolderPathValid;
    }

    /// <summary>
    /// Validates that the Excel file argument exists.
    /// </summary>
    private static bool ValidateExcelFile(string excelFilePath)
    {
        var excelFilePathValid = true;
        if (!File.Exists(excelFilePath))
        {
            Console.WriteLine($"Argument {nameof(excelFilePath)} error: '{excelFilePath}' was not found.");
            excelFilePathValid = false;
        }

        return excelFilePathValid;
    }

    /// <summary>
    /// Validates that the CsV folder argument exists and contains CSV files.
    /// </summary>
    private static bool ValidateCsvFolder(string csvFolderPath)
    {
        var folderExists = true;
        if (!Directory.Exists(csvFolderPath))
        {
            Console.WriteLine($"Argument {nameof(csvFolderPath)} error: '{csvFolderPath}' was not found.");
            folderExists = false;
        }

        var folderContainsCsvFiles = true;
        if (folderExists && !Directory.GetFiles(csvFolderPath, "*.csv").Any())
        {
            Console.WriteLine($"Argument {nameof(csvFolderPath)} error: No CSV files found in directory: {csvFolderPath}");
            folderContainsCsvFiles = false;
        }

        return folderExists && folderContainsCsvFiles;
    }
}