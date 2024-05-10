  ExcelLib Documentation

ExcelLib Documentation
======================

Overview
--------

`ExcelLib` is a C# library designed to facilitate reading from and writing to Excel files using the EPPlus library. It supports operations like reading data frames from Excel sheets and writing lists of objects back into Excel. The library aims to be easy to use while providing robust functionality for handling Excel data in .NET applications.

Features
--------

*   **Read Excel Data:** Load objects from Excel sheets into C# lists.
*   **Write to Excel:** Export lists of objects to new or existing Excel sheets.
*   **Dynamic Property Mapping:** Properties of objects are dynamically mapped to Excel columns based on custom attributes.
*   **Type Conversion and Validation:** Ensures that data conforms to expected types and formats, applying custom conversion rules as needed.

Installation
------------

To use `ExcelLib`, you must first ensure that the EPPlus package is installed in your project as it is a dependency for handling Excel file operations.

    Install-Package EPPlus -Version 5.x

Usage
-----

### Initializing the Library

    var excelLib = new ExcelLib("path/to/excel/file.xlsx");

### Reading from Excel

The `ReadDataFrame` method allows you to read data from a specified Excel sheet and map it to a list of objects of a specified type.

    var data = excelLib.ReadDataFrame<MyModel>("Sheet1");

### Writing to Excel

To write data to an Excel file, use the `WriteDataFrame` method. This method takes a list of objects and writes them to the specified Excel sheet.

    excelLib.WriteDataFrame(data, "Sheet1");

### Saving Changes

To save any changes made to the Excel file, use the `Save` or `SaveAs` method.

    excelLib.Save();  // Saves changes to the original file
    excelLib.SaveAs("path/to/new/file.xlsx");  // Saves changes to a new file

      ExcelLib Usage Examples

ExcelLib Usage Examples
=======================

Example Data Model Definitions
------------------------------

This section provides examples of defining models for inventory management and customer data processing.

### Inventory Item Example
```cs
    public class InventoryItem : ExcelDataModel {
        [Excel(Name = "Item ID", CanBeNull = false)]
        public string ItemId { get; set; }
    
        [Excel(Name = "Description", CaseSensitive = false)]
        public string Description { get; set; }
    
        [Excel(Name = "Quantity", Type = typeof(int))]
        public int Quantity { get; set; }
    
        [Excel(Name = "Price", Type = typeof(decimal))]
        public decimal Price { get; set; }
    }
            
```
### Customer Data Example

```cs
    public class Customer : ExcelDataModel {
        [Excel(Name = "Customer ID", CanBeNull = false)]
        public string CustomerId { get; set; }
    
        [Excel(Name = "Full Name")]
        public string FullName { get; set; }
    
        [Excel(Name = "Email Address", IgnoreCases = ["email", "mail"], CaseSensitive = false)]
        public string Email { get; set; }
    
        [Excel(Name = "Signup Date", Type = typeof(DateTime))]
        public DateTime SignupDate { get; set; }
    
        [Excel(Name = "Loyalty Points", Type = typeof(int), CanBeNull = true)]
        public int? LoyaltyPoints { get; set; }
    }
```    

Explanation of Attributes
-------------------------

### <a href="https://github.com/Lorefist5/Excel/tree/master/Excel.Library/Attributes">All attributes here</a>

Use Case: Reading Data
----------------------


To utilize these models, instantiate `ExcelLib`, specify the Excel file, and use the `ReadDataFrame` method with the model type that matches your Excel layout.
```cs
    var excelLib = new ExcelLib("path/to/inventory/file.xlsx");
    var inventoryItems = excelLib.ReadDataFrame<InventoryItem>("InventorySheet");
    var customerData = excelLib.ReadDataFrame<Customer>("Customers");
```
Error Handling
--------------

ExcelLib includes basic error handling capabilities, throwing exceptions when critical operations fail (such as file not found or sheet not existing).

Contributing
------------

Contributions to enhance ExcelLib, add features, or improve documentation are welcome. Please fork the repository and submit a pull request with your changes.

License
-------
