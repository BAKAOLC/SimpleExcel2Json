using System.Security.Cryptography;
using Newtonsoft.Json;

namespace Excel2Json;

public class FileMetaInfo
{
    public Dictionary<string, FileInfo> Files { get; set; } = new();

    public class FileInfo
    {
        public string InputPath { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;
        public string Hash { get; set; } = string.Empty;
        public DateTime LastModified { get; set; }
    }

    public static FileMetaInfo Load(string path)
    {
        if (!File.Exists(path))
            return new FileMetaInfo();

        var json = File.ReadAllText(path);
        return JsonConvert.DeserializeObject<FileMetaInfo>(json) ?? new FileMetaInfo();
    }

    public void Save(string path)
    {
        var json = JsonConvert.SerializeObject(this, Formatting.Indented);
        File.WriteAllText(path, json);
    }

    public static string CalculateFileHash(string filePath)
    {
        using var md5 = MD5.Create();
        using var stream = File.OpenRead(filePath);
        var hash = md5.ComputeHash(stream);
        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
    }
}