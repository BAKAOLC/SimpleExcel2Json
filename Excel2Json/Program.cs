using System.Text;
using CommandLine;
using Excel2Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;


ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var parserResult = Parser.Default.ParseArguments<Options>(args);
parserResult.WithParsed(BeginProcess);
parserResult.WithNotParsed(x =>
{
    Console.WriteLine("参数无效");
});

void BeginProcess(Options options)
{
    try
    {
        if (options.IsDirectoryMode)
        {
            ProcessDirectory(options);
        }
        else
        {
            ProcessSingleFile(options);
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"处理过程中发生错误: {ex.Message}");
        if (ex.InnerException != null)
        {
            Console.WriteLine($"内部错误: {ex.InnerException.Message}");
        }
    }
}

void ProcessDirectory(Options options)
{
    var inputDir = new DirectoryInfo(options.InputFile);
    if (!inputDir.Exists)
    {
        throw new DirectoryNotFoundException($"输入目录不存在: {inputDir.FullName}");
    }

    var outputDir = new DirectoryInfo(options.OutputFile);
    if (!outputDir.Exists)
    {
        outputDir.Create();
    }

    var metaFile = options.MetaFile;
    var metaInfo = FileMetaInfo.Load(metaFile);
    var excelFiles = inputDir.GetFiles("*.xlsx", SearchOption.AllDirectories);

    foreach (var excelFile in excelFiles)
    {
        try
        {
            var relativePath = PathHelper.GetRelativePath(inputDir.FullName, excelFile.FullName);
            var outputPath = Path.Combine(outputDir.FullName,
                Path.ChangeExtension(relativePath, ".json"));

            var fileInfo = new FileMetaInfo.FileInfo
            {
                InputPath = excelFile.FullName,
                OutputPath = outputPath,
                Hash = FileMetaInfo.CalculateFileHash(excelFile.FullName),
                LastModified = excelFile.LastWriteTime
            };

            var shouldProcess = !metaInfo.Files.TryGetValue(relativePath, out var existingInfo) ||
                              existingInfo.Hash != fileInfo.Hash ||
                              !File.Exists(outputPath) ||
                              File.GetLastWriteTime(outputPath) < excelFile.LastWriteTime;

            if (shouldProcess)
            {
                Console.WriteLine($"处理文件: {relativePath}");
                var fileOptions = new Options
                {
                    InputFile = excelFile.FullName,
                    OutputFile = outputPath,
                    BeginRow = options.BeginRow
                };

                ProcessSingleFile(fileOptions);
                metaInfo.Files[relativePath] = fileInfo;
            }
            else
            {
                Console.WriteLine($"跳过文件: {relativePath} (未更改)");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"处理文件 {excelFile.FullName} 时发生错误: {ex.Message}");
        }
    }

    metaInfo.Save(metaFile);
}

void ProcessSingleFile(Options options)
{
    if (!File.Exists(options.InputFile))
    {
        throw new FileNotFoundException("输入文件未找到", options.InputFile);
    }

    Console.WriteLine("读取文件: " + Path.GetFullPath(options.InputFile));
    using var package = new ExcelPackage(new FileInfo(options.InputFile));
    var worksheet = package.Workbook.Worksheets[0] ?? throw new InvalidOperationException("工作表中没有数据");
    Console.WriteLine("读取工作表: " + worksheet.Name);

    var (Success, Json, Errors) = ParseWorksheet(worksheet, options.BeginRow);

    if (Success)
    {
        Console.WriteLine("所有行解析成功");
        var outputPath = Path.GetFullPath(options.OutputFile);
        var outputDir = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }
        File.WriteAllText(outputPath, Json.ToString());
    }
    else
    {
        Console.WriteLine("解析过程中出现错误:");
        foreach (var error in Errors)
        {
            var errorType = error.IsWarning ? "WARNING" : "ERROR";
            Console.WriteLine($"[{errorType}] {options.InputFile}({error.Row},{error.Column}): {error.Message}");
        }
    }
}

(bool Success, JObject Json, List<ProcessError> Errors) ParseWorksheet(
    ExcelWorksheet worksheet, int beginRow)
{
    var resultJson = new JObject();
    var jsonArray = new JArray();
    resultJson["data"] = jsonArray;

    if (worksheet.Dimension == null)
    {
        return (false, resultJson, new List<ProcessError>
        {
            new("工作表为空", 0, 0, false)
        });
    }

    var headers = new string[worksheet.Dimension.Columns];
    var recordExceptions = new List<ProcessError>();
    var success = true;

    for (var row = beginRow; row <= worksheet.Dimension.Rows; row++)
    {
        if (row == beginRow)
        {
            for (var headCol = 1; headCol <= worksheet.Dimension.Columns; headCol++)
            {
                var cell = worksheet.Cells[row, headCol];
                headers[headCol - 1] = cell?.Text ?? $"Column{headCol}";
            }
            continue;
        }

        var data = new JObject();
        for (var col = 1; col <= worksheet.Dimension.Columns; col++)
        {
            try
            {
                var cell = worksheet.Cells[row, col];
                data[headers[col - 1]] = ValueHelper.CheckValue(cell?.Text ?? string.Empty);
            }
            catch (Exception ex)
            {
                var sb = new StringBuilder();
                sb.AppendLine($"行 {row} 列 {col}: {worksheet.Cells[row, col]?.Text ?? "空"}");
                sb.AppendLine(ex.Message);
                recordExceptions.Add(new ProcessError(sb.ToString(), row, col, false));
                success = false;
            }
        }

        if (success)
        {
            jsonArray.Add(data);
        }
    }

    return (success, resultJson, recordExceptions);
}