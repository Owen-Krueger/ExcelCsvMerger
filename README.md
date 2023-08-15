# Excel CSV Merger
Excel CSV Manager is a CLI application for merging Excel and CSV documents.

## Prepping Files
In order to use Excel CSV Merger, you need two things:
- The Excel document to replicate
- A folder of CSV files formatted correctly

Now, what do I mean by formatted correctly? I mean that the file and column names match perfectly to data elements on the Excel document. A single CSV file should represent all the data to merge into an Excel sheet on the Excel document with the same name. Each header column on that CSV file should match the header column on the Excel document.

These CSV files also do not need to have every column from the FBDI document. The program will skip over any columns in the FBDI document that don’t exist in the CSV.

## Running
The program takes in two arguments and an optional option. You can also execute .\ExcelCsvMerger.exe -h to get help on what the program needs.

| Name | Required | Description |
| --- | :---: | --- |
| excelFilePath | Yes | File path to Excel document to merge CSV files into. |
| csvFolderPath | Yes | Path to folder with CSV files to merge in. Folder must contain 'csv' files. |
| --output | No | Optional. Specify file path to export merged FBDI document to. |

If `--output` is not provided, the output file will be generated in the same folder as the FBDI document with the same extension and will be named Output.