using System.Runtime.InteropServices;
using MiniExcelLibs;
using Excel = Microsoft.Office.Interop.Excel;

namespace JbFileProcessor.Core;

public static class FileUtils
{
	/// <summary>
	/// The file extensions that are valid for excel files
	/// </summary>
	public static readonly string[] ValidExcelExtensions = {".xlsx", ".xls"};
	/// <summary>
	/// The file extension that is valid for csv files
	/// </summary>
	public static readonly string ValidCsvExtension = ".csv";

	// This method will read a source text file line by line, replaces template tags with values and writes the result to a destination file.
	public static void ProcessFile(string sourceFilePath, string destinationFilePath,
		Dictionary<string, string> templateTags)
	{
		// Read the source file line by line
		var lines = File.ReadAllLines(sourceFilePath);

		// Replace template tags with values
		for (var i = 0; i < lines.Length; i++)
		{
			foreach (var tag in templateTags)
			{
				lines[i] = lines[i].Replace($"{{{{{tag.Key}}}}}", tag.Value);
			}
		}

		// Write the result to the destination file
		File.WriteAllLines(destinationFilePath, lines);
	}
	
	/// <summary>
	/// Read a csv file, and return a list of dictionaries.
	/// Each dictionary represents a row in the csv file.
	/// </summary>
	/// <param name="filePath">The file to read from</param>
	/// <param name="delimiter">The char that is used to delimit the columns</param>
	/// <returns>The imported data</returns>
	public static IEnumerable<Dictionary<string, string>> ReadCsvFile(string filePath, char delimiter = ',')
	{
		if (!File.Exists(filePath))
			throw new FileNotFoundException("The file does not exist", filePath);
		
		var result = new List<Dictionary<string, string>>();

		// Read the csv file line by line
		var lines = File.ReadAllLines(filePath);

		// Get the header row
		var headerRow = lines[0];

		// Get the header columns
		var headerColumns = headerRow.Split(delimiter);

		// Loop through the data rows
		for (var i = 1; i < lines.Length; i++)
		{
			// Get the data row
			var dataRow = lines[i];

			// Get the data columns
			var dataColumns = dataRow.Split(delimiter);

			// Create a dictionary to hold the data
			var data = new Dictionary<string, string>();

			// Loop through the header columns and add the data to the dictionary
			for (var j = 0; j < headerColumns.Length; j++)
			{
				data.Add(headerColumns[j], dataColumns[j]);
			}

			// Add the dictionary to the result
			result.Add(data);
		}

		return result;
	}

	public static IEnumerable<Dictionary<string, string>> ReadExcelFile(string filePath)
	{
		if (!File.Exists(filePath))
			throw new FileNotFoundException("The file does not exist", filePath);
		
		var result = new List<Dictionary<string, string>>();

		// Get the header columns
		var headerColumns = ((IDictionary<string, object>)MiniExcel.Query(filePath, useHeaderRow: true).First()).Keys.ToList();
		
		foreach(IDictionary<string,object> row in MiniExcel.Query(filePath).Skip(1))
		{
			var data = new Dictionary<string, string>();
			
			// Columns
			for (var i = 0; i < row.Count; i++)
			{
				var value = row.Values.ElementAt(i);
				data.Add(headerColumns[i], value.ToString() ?? "[ERROR]");
			}

			result.Add(data);
		}

		return result;
	}
	
	/// <summary>
	/// Convert a excel file into a new csv file
	/// </summary>
	/// <param name="excelSourceFilePath">The source file that contains the excel data</param>
	/// <param name="csvDestinationFilePath">The destination file that will contain the csv data</param>
	/// <exception cref="Exception">If the file could not be converted. See the inner exception for further information</exception>
	public static void ConvertExcelFileToCsv(string excelSourceFilePath, string csvDestinationFilePath)
	{
		Excel.Application? excelApp = null;
		Excel.Workbook? excelWorkbook = null;

		try
		{
			// Create a new excel application
			excelApp = new Excel.Application();

			if (excelApp == null)
				throw new Exception("Excel is not properly installed!!");
		
			// check that file exist
			if (!File.Exists(excelSourceFilePath))
				throw new FileNotFoundException("The file does not exist", excelSourceFilePath);
		
			// Check if the file has a valid extension
			if (!HasValidFileExtension(excelSourceFilePath, ValidExcelExtensions))
				throw new Exception("The file is not a valid excel file");
		
			// Open the excel file
			excelWorkbook = excelApp.Workbooks.Open(excelSourceFilePath);
		
			if (excelWorkbook == null)
				throw new Exception("The file could not be opened");

			var worksheet = (Excel.Worksheet) excelApp.ActiveWorkbook.Sheets[1];

			if (worksheet == null)
				throw new Exception("The worksheet could not be opened");
		
			// // Get the used range
			// var usedRange = worksheet.UsedRange;
		
			worksheet.SaveAs(csvDestinationFilePath, Excel.XlFileFormat.xlCSV);
		}
		catch (Exception e)
		{
			throw new Exception("Could not convert excel file to csv", e);
		}
		finally
		{
			excelWorkbook.Close();
			excelApp.Quit();
		}
	}

	
	/// <summary>
	/// Check if a file has a extension, from a list of valid extensions
	/// </summary>
	/// <param name="filePath">The path to check</param>
	/// <param name="validExtensions">The list of valid extensions</param>
	/// <returns></returns>
	public static bool HasValidFileExtension(string filePath, IEnumerable<string> validExtensions)
	{
		// Get the file extension
		var fileExtension = Path.GetExtension(filePath);

		// Check if the file extension is valid
		return validExtensions.Contains(fileExtension);
	}
	
	/// <summary>
	/// Get a unique file path from a source file path
	/// </summary>
	/// <param name="sourceFilePath"></param>
	/// <returns></returns>
	public static string GetUniquePath(string sourceFilePath)
	{
		// Get the file name
		var fileName = Path.GetFileName(sourceFilePath);

		// Get the file extension
		var fileExtension = Path.GetExtension(sourceFilePath);

		// Get the file name without the extension
		var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(sourceFilePath);

		// Get the directory path
		var directoryPath = Path.GetDirectoryName(sourceFilePath);

		// Get the unique file name
		var uniqueFileName = $"{fileNameWithoutExtension}_{Guid.NewGuid()}{fileExtension}";

		// Get the unique file path
		var uniqueFilePath = Path.Combine(directoryPath, uniqueFileName);

		return uniqueFilePath;
	}
}