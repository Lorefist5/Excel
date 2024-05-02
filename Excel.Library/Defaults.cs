using Excel.Library.Enums;

namespace Excel.Library;

public class Defaults
{
    public static bool DefaultCanBeNullValue { get; set; }
    public static TrimMode DefaultTrimMode { get; set; } = TrimMode.FrontAndEnd;
}
