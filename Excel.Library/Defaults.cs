using Excel.Library.Enums;

namespace Excel.Library;

public class Defaults
{
    public static bool DefaultCanBeNullValue { get; set; }
    public static TrimMode DefaultTrimMode { get; set; }
    private static int _ignoreHeaderCount;
    private static int _ignoreLastRowCount;
    public static int IgnoreHeaderCount
    {
        get => _ignoreHeaderCount;
        set => _ignoreHeaderCount = value <= 0 ? 1 : value;
    }
    public static int IgnoreLastRowCount
    {
        get => _ignoreLastRowCount;
        set => _ignoreLastRowCount = value <= 0 ? 1 : value;
    }
}
