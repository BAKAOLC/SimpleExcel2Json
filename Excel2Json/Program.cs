﻿using System.Text;
using CommandLine;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var parserResult = Parser.Default.ParseArguments<Options>(args);
await parserResult.WithParsedAsync(BeginParseExcel).ConfigureAwait(false);
await parserResult.WithNotParsedAsync(x =>
{
    Console.WriteLine("Invalid arguments");
    return Task.CompletedTask;
}).ConfigureAwait(false);

return;

Task BeginParseExcel(Options options)
{
    if (!File.Exists(options.InputFile))
    {
        Console.WriteLine("Input file not found: " + Path.GetFullPath(options.InputFile));
        return Task.CompletedTask;
    }

    Console.WriteLine("Reading file: " + Path.GetFullPath(options.InputFile));
    using var package = new ExcelPackage(options.InputFile);
    Console.WriteLine("Reading worksheet: " + package.Workbook.Worksheets[0].Name);
    var worksheet = package.Workbook.Worksheets[0];
    var beginRow = options.BeginRow;
    var resultJson = new JObject();
    var jsonArray = new JArray();
    resultJson["data"] = jsonArray;
    var headers = new string[worksheet.Dimension.Columns];

    var recordExceptions = new List<string>();
    var success = true;

    for (var row = beginRow; row <= worksheet.Dimension.Rows; row++)
    {
        if (row == beginRow)
        {
            for (var headCol = 1; headCol <= worksheet.Dimension.Columns; headCol++)
                headers[headCol - 1] = worksheet.Cells[row, headCol].Text;

            Console.WriteLine("Headers: " + string.Join(", ", headers));
            continue;
        }

        if (!success)
        {
            for (var col = 1; col <= worksheet.Dimension.Columns; col++)
                try
                {
                    CheckValue(worksheet.Cells[row, col].Text);
                }
                catch (Exception ex)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine($"Row {row} Column {col}: {worksheet.Cells[row, col].Text}");
                    sb.AppendLine(ex.Message);
                    recordExceptions.Add(sb.ToString());
                }

            continue;
        }

        var data = new JObject();
        for (var col = 1; col <= worksheet.Dimension.Columns; col++)
        {
            if (!success)
            {
                try
                {
                    CheckValue(worksheet.Cells[row, col].Text);
                }
                catch (Exception ex)
                {
                    var sb = new StringBuilder();
                    sb.AppendLine($"Row {row} Column {col}: {worksheet.Cells[row, col].Text}");
                    sb.AppendLine(ex.Message);
                    recordExceptions.Add(sb.ToString());
                }

                continue;
            }

            try
            {
                data[headers[col - 1]] = CheckValue(worksheet.Cells[row, col].Text);
            }
            catch (Exception ex)
            {
                var sb = new StringBuilder();
                sb.AppendLine($"Row {row} Column {col}: {worksheet.Cells[row, col].Text}");
                sb.AppendLine(ex.Message);
                recordExceptions.Add(sb.ToString());
                success = false;
            }
        }

        Console.WriteLine($"Row {row - beginRow + 1}: " + data);

        if (success) jsonArray.Add(data);
    }

    if (success)
    {
        Console.WriteLine("All rows parsed successfully.");
        Console.WriteLine("Writing to file: " + Path.GetFullPath(options.OutputFile));
        File.WriteAllText(options.OutputFile, resultJson.ToString());
    }
    else
    {
        Console.WriteLine("Errors occurred while parsing the following rows:");
        foreach (var recordException in recordExceptions)
            Console.WriteLine(recordException);
    }

    return Task.CompletedTask;
}

JToken CheckValue(string value)
{
    if (value.StartsWith("{") || value.StartsWith("["))
        return JToken.Parse(value);

    if (int.TryParse(value, out var intValue))
        return intValue;

    if (double.TryParse(value, out var doubleValue))
        return doubleValue;

    return value;
}

internal class Options
{
    [Option('i', "input", Required = true, HelpText = "Input Excel file")]
    public required string InputFile { get; set; }

    [Option('o', "output", Required = true, HelpText = "Output JSON file")]
    public required string OutputFile { get; set; }

    [Option('r', "row", Required = false, HelpText = "Begin row to read data")]
    public int BeginRow { get; set; } = 2;
}