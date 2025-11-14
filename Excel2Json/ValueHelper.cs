using Newtonsoft.Json.Linq;

namespace Excel2Json;

public static class ValueHelper
{
    public static JToken CheckValue(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        if (value.StartsWith("{") || value.StartsWith("["))
        {
            try
            {
                return JToken.Parse(value);
            }
            catch
            {
                return value;
            }
        }

        if (int.TryParse(value, out var intValue))
            return intValue;

        if (double.TryParse(value, out var doubleValue))
            return doubleValue;

        return value;
    }
} 