namespace Excel2Json;

public static class PathHelper
{
    public static string GetRelativePath(string basePath, string fullPath)
    {
        try
        {
            if (!basePath.EndsWith(Path.DirectorySeparatorChar.ToString()))
            {
                basePath += Path.DirectorySeparatorChar;
            }

            var normalizedBasePath = Path.GetFullPath(basePath);
            var normalizedFullPath = Path.GetFullPath(fullPath);

            if (!normalizedFullPath.StartsWith(normalizedBasePath, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("文件不在指定的基础目录内");
            }

            var relativePath = normalizedFullPath.Substring(normalizedBasePath.Length);

            return relativePath.Replace('\\', Path.DirectorySeparatorChar)
                              .Replace('/', Path.DirectorySeparatorChar);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"无法获取相对路径: {ex.Message}", ex);
        }
    }
}