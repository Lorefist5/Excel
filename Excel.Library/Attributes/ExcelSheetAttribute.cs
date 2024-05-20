namespace Excel.Library.Attributes;
[AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
public class ExcelSheetAttribute : Attribute
{
    public string Name { get; set; } // Read and write property for the Sheet
    public string[]? ReadingProperties { get; set; } // When searching for the sheet it will also check for all these names
    public bool ReadMultiple { get; set; } // If it will read multiple sheets if not it will stop at the first sheet that it finds
    public int? Index { get; set; } = null; // The index of the sheet that will be read this way it will only read the sheet in that index only
}
