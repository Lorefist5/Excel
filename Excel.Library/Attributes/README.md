# ExcelAttribute Configuration Guide

## Attribute Properties Explanation

This section describes each property available in the `ExcelAttribute` class, providing examples of their application for data models.

### **Name**
- Specifies the exact column name in the Excel sheet that maps to the C# property.

### **IsReadProperty**
- Determines if the property should be read from the Excel file (default is true).

### **IsWriteProperty**
- Indicates whether the property should be written to the Excel file when exporting data (default is true).

### **IndexOrder**
- Sets the order in which properties are written to the Excel file when exporting data. Lower numbers are written first.
## **Index of header**
- Specifies the index of the header in the Excel file. This is useful when the header is in the same column in each excel you read.
### **DefaultValue**
- Specifies a default value to use when the Excel cell is empty or null during import.

### **CaseSensitive**
- Controls whether the Excel column names are case-sensitive when matching with property names (default is true).

### **IgnoreCases**
- Specifies substrings to be removed from the Excel cell's value during import.

### **ReadingProperties**
- Allows multiple column names to be associated with a single property. Useful for accommodating variations in column headers.

### **CaseStyle**
- Dictates the case style to apply when writing data to Excel.

### **IgnoreHeaderCases**
- Substrings to be ignored in header names during import, helping to streamline data consistency.

### **TrimMode**
- Specifies how whitespace should be trimmed from Excel cell values during import.

### **CanBeNull**
- Indicates if the property can accept null values, affecting data validation (default is false).

### **Type**
- The data type to which the Excel cell value should be converted upon import.

## CaseStyle Enum

Defines the case formatting to apply to string values when written to an Excel file.

- **CamelCase:** firstLetterLowerCaseSubsequentWordsCapitalized
- **SnakeCase:** words_separated_by_underscores
- **PascalCase:** EachWordCapitalized
- **Lower:** all letters in lower case
- **Upper:** ALL LETTERS IN UPPER CASE
- **Default:** No change to the original text.

## TrimMode Enum

Describes how whitespace is managed for Excel cell values during data import.

- **End:** Trims whitespace from the end of the string.
- **Front:** Trims whitespace from the beginning of the string.
- **FrontAndEnd:** Trims whitespace from both ends of the string.
- **All:** Removes all whitespace from the string.
