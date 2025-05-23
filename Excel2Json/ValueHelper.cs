using Newtonsoft.Json.Linq;

namespace Excel2Json;

public static class ValueHelper
{
    public static JToken CheckValue(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        if (BBCodeHelper.IsBBCode(value))
            return value;

        if (value.StartsWith("{") || value.StartsWith("["))
        {
            try
            {
                return JToken.Parse(value);
            }
            catch
            {
                throw new ProcessWarningException("JSON解析失败，将作为普通字符串处理");
            }
        }

        if (int.TryParse(value, out var intValue))
            return intValue;

        if (double.TryParse(value, out var doubleValue))
            return doubleValue;

        return value;
    }
} 