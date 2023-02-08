

using JbFileProcessor.Core;

Console.WriteLine("Select a excel file to process");

// Get the file path
// var filePath = Console.ReadLine();
const string excelFile = "DataTest.xlsx";
const string csvFile = "DataTest.csv";
const string templateFile = "Template.txt";
const string targetFile = "Result.txt";

// Check if the file exists
if (!File.Exists(excelFile))
{
	Console.WriteLine("The file does not exist");
	return;
}

// Check if the file has a valid extension
if (!FileUtils.HasValidFileExtension(excelFile, FileUtils.ValidExcelExtensions))
{
	Console.WriteLine("The file is not a valid excel file");
	return;
}

try
{
	var templateData = FileUtils.ReadExcelFile(excelFile);
	
	var fileProcessor = new FileProcessor(new FileProcessorOptions()
	{
		TemplateFile = "Template.txt",
		TemplateData = templateData,
		GetDestinationFilePathFromTemplateData = true
	});

	var files = await fileProcessor.Process();

	Console.WriteLine(files.Count == 1 ? $"Created {files.Count} file" : $"Created {files.Count} files");
	
	foreach (var file in files)
	{
		Console.WriteLine($" - {file}");
	}
}
catch (Exception e)
{
	Console.WriteLine("Error !");
	Console.WriteLine(e);
}