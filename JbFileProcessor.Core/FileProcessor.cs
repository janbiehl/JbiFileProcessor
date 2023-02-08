namespace JbFileProcessor.Core;

public struct FileProcessorOptions
{
	public FileProcessorOptions()
	{
	}

	/// <summary>
	/// Contains the template tags and their values
	/// </summary>
	public required IEnumerable<Dictionary<string, string>> TemplateData { get; init; }

	/// <summary>
	/// The file that contains the templates to be replaced
	/// </summary>
	public required string TemplateFile { get; init; }

	/// <summary>
	/// When this flag is activated, the destination file path will be read from the template data.
	/// </summary>
	public bool GetDestinationFilePathFromTemplateData { get; init; } = false;
}

public class FileProcessor
{
	private readonly FileProcessorOptions _options;

	public FileProcessor(FileProcessorOptions options)
	{
		_options = options;
	}

	public async Task<List<string>> Process(CancellationToken cancellationToken = default)
	{
		if (!File.Exists(_options.TemplateFile))
			throw new FileNotFoundException("The template file does not exist", _options.TemplateFile);
		
		List<string> destinationFiles = new();

		foreach (var templateFileData in _options.TemplateData)
		{
			// Get the destination file path from the template data
			var destinationFileName = _options.GetDestinationFilePathFromTemplateData ? 
				GetDestinationFileNameFromTemplateData(templateFileData) 
				: Path.GetFileNameWithoutExtension(_options.TemplateFile);

			var destinationFilePath = GetDestinationFilePath(_options.TemplateFile, null, destinationFileName);
			
			await ProcessFile(_options.TemplateFile, destinationFilePath, templateFileData, cancellationToken);
			destinationFiles.Add(destinationFilePath);
		}

		return destinationFiles;
	}

	private static string GetDestinationFilePath(string sourceFilePath, string? directory, string destinationFileName)
	{
		if (string.IsNullOrWhiteSpace(sourceFilePath))
			throw new ArgumentException("Value cannot be null or whitespace.", nameof(sourceFilePath));

		var fileExtension = Path.GetExtension(sourceFilePath);
		
		var filePathWithoutExtension = directory is null ?
			// Use the same directory as the source file
			Path.Combine(Path.GetDirectoryName(sourceFilePath)!, destinationFileName) :
			// Use the specified directory
			Path.Combine(directory, destinationFileName);

		return Path.ChangeExtension(filePathWithoutExtension, fileExtension);
	}

	/// <summary>
	/// Get the destination file path from the template data.
	/// </summary>
	/// <param name="templateFileData">The data that will be used to process the file</param>
	/// <returns>The file name for the destination file</returns>
	/// <exception cref="Exception">The template data does not contain the desired keys</exception>
	private static string GetDestinationFileNameFromTemplateData(IReadOnlyDictionary<string, string> templateFileData)
	{
		string? destinationFileName = null;
		
		if (templateFileData.ContainsKey("FileName"))
		{
			destinationFileName = templateFileData["FileName"];
		}
		else if (templateFileData.ContainsKey("Dateiname"))
		{
			destinationFileName = templateFileData["Dateiname"];
		}
		
		if (destinationFileName == null)
		{
			throw new Exception("The template data does not contain a 'FileName' or 'Dateiname' key");
		}
		
		return destinationFileName;
	}

	/// <summary>
	/// Read the source file line by line, replaces template tags with values and writes the result to a destination file.
	/// </summary>
	/// <param name="sourceFile">The file that contains the templates</param>
	/// <param name="destinationFile">The path for the created file</param>
	/// <param name="templateFileData">The values the templates will be replaced with</param>
	/// <param name="cancellationToken">Used to cancel the operation</param>
	/// <returns>The path to the newly created file</returns>
	private static async Task<string> ProcessFile(string sourceFile, string destinationFile, Dictionary<string, string> templateFileData, CancellationToken cancellationToken = default)
	{
		using var sourceStream = File.OpenText(sourceFile);
		await using var destinationStream = File.CreateText(destinationFile);

		while (await sourceStream.ReadLineAsync(cancellationToken) is { } sourceLine)
		{
			// Replace template tags with values
			var destinationLine = templateFileData
				.Aggregate(sourceLine, (current, tag) 
					=> current.Replace($"{{{{{tag.Key}}}}}", tag.Value));

			// Write the result to the destination file
			await destinationStream.WriteLineAsync(destinationLine);
		}

		return destinationFile;
	}
}