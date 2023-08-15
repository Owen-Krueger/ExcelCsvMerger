using System.Globalization;
using CsvHelper;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelCsvMerger;

/// <summary>
/// For generating Excel documents.
/// </summary>
public static class ExcelMerger
{
    /// <summary>
    /// Merges CSV files into output xlsm file.
    /// </summary>
    /// <param name="outputFilePath">File path to file being merged into.</param>
    /// <param name="csvFolder">Folder containing CSV files.</param>
    public static void MergeFiles(string outputFilePath, string csvFolder)
    {
        using var spreadsheetDocument = SpreadsheetDocument.Open(outputFilePath, true);
        var workbook = spreadsheetDocument.WorkbookPart;
            
        foreach (var filePath in Directory.GetFiles(csvFolder, "*.csv"))
        {
            try
            {
                var records = GetCsvRecords(filePath);
                var sheetData = GetWorksheetData(workbook, filePath);

                if (sheetData == null)
                {
                    continue;
                }
                
                var headerRow = sheetData.Elements<Row>().First();
                var headers = headerRow.Elements<Cell>().Select(cell => GetCellValue(workbook, cell)).ToList();

                var progressBar = new ProgressBar($"Merging {Path.GetFileNameWithoutExtension(filePath)}", records.Count);
                foreach (var record in records)
                {
                    progressBar.IncrementProgressBar();
                    AddRow(sheetData, record, headers);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Exception raised while merging in {Path.GetFileNameWithoutExtension(filePath)} file: {e.Message}");
            }
        }

        workbook.Workbook.Save();
    }

    /// <summary>
    /// Gets all records from CSV file at input filepath.
    /// </summary>
    /// <param name="filePath">Filepath containing CSV files.</param>
    /// <returns>Dynamic records.</returns>
    private static List<dynamic> GetCsvRecords(string filePath)
    {
        using var reader = new StreamReader(filePath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        return csv.GetRecords<dynamic>().ToList();
    }

    /// <summary>
    /// Gets worksheet data for sheet matching CSV file name.
    /// </summary>
    /// <param name="workbook">Workbook containing sheets.</param>
    /// <param name="filePath">Path to CSV file.</param>
    /// <returns>Worksheet data.</returns>
    private static SheetData? GetWorksheetData(WorkbookPart workbook, string filePath)
    {
        var sheetName = Path.GetFileNameWithoutExtension(filePath);
        var worksheet = FindWorksheet(workbook, sheetName);

        if (worksheet == null)
        {
            Console.WriteLine($"{sheetName} does not match a sheet in FBDI document. Skipping.");
            return null;
        }
        
        return worksheet.Worksheet.Elements<SheetData>().First();
    }

    /// <summary>
    /// Adds a new row to the sheet.
    /// </summary>
    /// <param name="sheetData">Worksheet to add row to.</param>
    /// <param name="record">CSV record to add.</param>
    /// <param name="headers">Headers to add contents under.</param>
    private static void AddRow(OpenXmlElement sheetData, dynamic record, List<string> headers)
    {
        var newRow = new Row();

        foreach (var header in headers)
        {
            var newCell = new Cell
            {
                DataType = CellValues.String
            };

            if (TryGetValueFromCsv(record, header, out object value))
            {
                newCell.CellValue = new CellValue(value.ToString());
            }

            newRow.AppendChild(newCell);
        }
                
        sheetData.AppendChild(newRow);
    }

    /// <summary>
    /// Trys to get value from CSV under <see cref="memberName"/> column.
    /// </summary>
    /// <param name="obj">Row in CSV.</param>
    /// <param name="memberName">Column name.</param>
    /// <param name="value">Value from CSV, if found.</param>
    /// <returns>True if value found in CSV.</returns>
    private static bool TryGetValueFromCsv(dynamic obj, string memberName, out object? value)
    {
        return ((IDictionary<string, object>)obj).TryGetValue(memberName, out value);
    }

    /// <summary>
    /// Finds worksheet by name.
    /// </summary>
    /// <param name="workbookPart">Workbook containing worksheet.</param>
    /// <param name="sheetName">Sheet name to search for.</param>
    /// <returns>The worksheet, if found. Null if not found.</returns>
    private static WorksheetPart? FindWorksheet(WorkbookPart workbookPart, string sheetName)
    {
        var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name.Value == sheetName);
        return (sheet != null) ? workbookPart.GetPartById(sheet.Id) as WorksheetPart : null;
    }

    /// <summary>
    /// Gets the string contents of the cell.
    /// </summary>
    /// <param name="workbookPart">Worksheet containing the cell.</param>
    /// <param name="cell">Cell to pull text from.</param>
    /// <returns>String contents of the input cell.</returns>
    private static string GetCellValue(WorkbookPart workbookPart, CellType cell)
    {
        var value = cell.CellValue.InnerText;

        if (cell.DataType is not { Value: CellValues.SharedString })
        {
            return value;
        }
        
        var index = int.Parse(value);
        var ssi = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(index);
        value = ssi.Text.Text;

        return value;
    }
}

