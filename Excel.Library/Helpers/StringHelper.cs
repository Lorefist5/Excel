using Excel.Library.Attributes;
using Excel.Library.Enums;
using Microsoft.Extensions.Primitives;

namespace Excel.Library.Helpers;

public static class StringHelper
{
    public static string RemoveIgnoreCases(string value,IEnumerable<string> ignoreCases, bool caseSensitive)
    {
        foreach (var ignoreCase in ignoreCases)
        {
            if (value.StartsWith(ignoreCase, caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase))
            {
                value = value.Remove(0, ignoreCase.Length).Trim();
            }
        }
        return value;
    }
    public static string RemoveSubstrings(this string mainString, IEnumerable<string> substringsToRemove)
    {
        foreach (var substring in substringsToRemove)
        {
            mainString = mainString.Replace(substring, "");
        }

        return mainString;
    }
    public static string ConvertToCaseStyle(string input, CaseStyle caseStyle, bool removeWhiteSpace = false)
    {
        if (removeWhiteSpace)
        {
            input = input.Replace(" ", string.Empty);
        }

        switch (caseStyle)
        {
            case CaseStyle.CamelCase:
                return ConvertToCamelCase(input);
            case CaseStyle.SnakeCase:
                return ConvertToSnakeCase(input);
            case CaseStyle.PascalCase:
                return ConvertToPascalCase(input);
            case CaseStyle.Lower:
                return input.ToLower();
            case CaseStyle.Upper:
                return input.ToUpper();
            case CaseStyle.Default:
                return input;
            default:
                throw new ArgumentOutOfRangeException(nameof(caseStyle), caseStyle, null);
        }
    }

    private static string ConvertToCamelCase(string input)
    {
        if (string.IsNullOrEmpty(input)) return input;
        var modifiedInput = ConvertToPascalCase(input); // Convert to PascalCase first to capitalize the first letter of each word.
        return char.ToLowerInvariant(modifiedInput[0]) + modifiedInput.Substring(1);
    }

    private static string ConvertToPascalCase(string input)
    {
        if (string.IsNullOrEmpty(input)) return input;
        // Split the string into words, capitalize the first letter of each word, and concatenate them.
        return string.Join("", input.Split(' ').Select(word => char.ToUpperInvariant(word[0]) + word.Substring(1).ToLower()));
    }

    private static string ConvertToSnakeCase(string input)
    {
        if (string.IsNullOrEmpty(input)) return input;
        // Insert underscores before each uppercase letter (except the first one) and convert the entire string to lowercase.
        return string.Concat(input.Select((x, i) => i > 0 && char.IsUpper(x) ? "_" + x.ToString().ToLower() : x.ToString().ToLower()));
    }
}
