using CommandLine;

namespace Excel2Json;

public class Options
{
    [Option('i', "input", Required = false, HelpText = "输入Excel文件或文件夹路径，默认为当前目录下的excel文件夹")]
    public string InputFile { get; set; } = Path.Combine(Directory.GetCurrentDirectory(), "excel");

    [Option('o', "output", Required = false, HelpText = "输出JSON文件或文件夹路径，默认为当前目录下的json文件夹")]
    public string OutputFile { get; set; } = Path.Combine(Directory.GetCurrentDirectory(), "json");

    [Option('m', "meta", Required = false, HelpText = "元数据文件路径，默认为当前目录下的.e2jmeta文件")]
    public string MetaFile { get; set; } = Path.Combine(Directory.GetCurrentDirectory(), ".e2jmeta");

    [Option('r', "row", Required = false, HelpText = "开始读取数据的行号，默认为2")]
    public int BeginRow { get; set; } = 2;

    public bool IsDirectoryMode => Directory.Exists(InputFile);
}