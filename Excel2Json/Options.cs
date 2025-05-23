using CommandLine;

namespace Excel2Json;

public class Options
{
    [Option('i', "input", Required = true, HelpText = "输入Excel文件或文件夹路径")]
    public required string InputFile { get; set; }

    [Option('o', "output", Required = true, HelpText = "输出JSON文件或文件夹路径")]
    public required string OutputFile { get; set; }

    [Option('m', "meta", Required = false, HelpText = "元数据文件路径")]
    public string? MetaFile { get; set; }

    [Option('r', "row", Required = false, HelpText = "开始读取数据的行号")]
    public int BeginRow { get; set; } = 2;

    public bool IsDirectoryMode => Directory.Exists(InputFile);
}