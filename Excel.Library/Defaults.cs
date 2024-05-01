using Excel.Library.Enums;

namespace Excel.Library;

public class Defaults
{
    public static bool DefaultCanBeNullValue { get; set; }
    public static TrimMode DefaultTrimMode { get; set; }
    public static int IgnoreHeaderCount { get; set; } = 0; //if the amount of null headers "ReadDataFrame" can count until it decides that its not part of the dataframe
}
