namespace Excel2Json;

public static class BBCodeHelper
{
    private static readonly string[] BBTags = new[]
    {
        "[b]", "[/b]", "[i]", "[/i]", "[u]", "[/u]", "[s]", "[/s]",
        "[url]", "[/url]", "[img]", "[/img]", "[code]", "[/code]",
        "[quote]", "[/quote]", "[list]", "[/list]", "[*]", "[color]",
        "[/color]", "[size]", "[/size]", "[font]", "[/font]", "[align]",
        "[/align]", "[table]", "[/table]", "[tr]", "[/tr]", "[td]", "[/td]"
    };

    public static bool IsBBCode(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return false;

        if (!value.StartsWith("[", StringComparison.OrdinalIgnoreCase))
            return false;

        return BBTags.Any(tag => value.StartsWith(tag, StringComparison.OrdinalIgnoreCase));
    }
}