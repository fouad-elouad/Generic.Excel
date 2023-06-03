# Generic.Excel

A .NET class library that provides a convenient and flexible way to generate .xlsx files for any .NET type. With the ability to customize property export, you can easily export data to Excel with ease.

## Description

Generic.Excel is a .NET class library based on the ClosedXML library. It provides a generic abstraction for generating XLSX files for any .NET type.
The library uses the ExcelPropertyAttribute to decorate model properties, allowing you to export properties with specified display names, order, 
and whether they should be ignored or not. It also supports nested properties for referencing other properties.

## Features
- [x] Target .NET Standard 2.0
- [x] Generate XLSX files for any .NET type.
- [x] Decorate model properties with ExcelPropertyAttribute for export customization.
- [x] Specify display names, order, and ignored properties using attributes.
- [x] Support for nested properties.
- [x] Explicit definition of properties to export.
- [x] Control the direction of displayed data (vertical or horizontal).
- [x] Unit Test Project

With these features, Generic.Excel simplifies the process of generating XLSX files in .NET applications by providing a generic and customizable approach. 
It gives you control over the exported data and allows you to easily work with complex models and relationships.

## Overview

### Generic Abstraction
Generic.Excel provides a generic abstraction for generating XLSX files for any .NET type. 
This means you can easily export data from different types of models without writing specific export logic for each type.

### ExcelPropertyAttribute
The library utilizes the ExcelPropertyAttribute, which can be applied to model properties. 
This attribute allows you to customize the export behavior of properties by specifying display names, order, and whether they should be ignored or not.
This gives you fine-grained control over the exported data.

The attribute provides the following properties:

    - DisplayName: Specifies the display name of the property in the Excel file.
    - Ignore: Indicates whether the property should be ignored during export.
    - Order: Specifies the order of the property in the Excel file.
    - NestedProperty: Specifies a nested property to be used for reference properties.
    
### Nested Properties
Generic.Excel supports nested properties, allowing you to reference and export properties from related models. 
This is useful when you have complex data structures or relationships between models and want to export them in a structured manner.

### Explicit definition of properties to export

With Generic.Excel, you have control over which properties are included in the export process.
This allows you to selectively choose the properties you want to export, providing flexibility and customization options based on your specific requirements.

### Export Direction
The library gives you the ability to control the direction in which data is displayed in the Excel file.
By default, data is exported vertically. You can choose between vertical orientation (where each object is displayed in a separate row) or 
horizontal orientation (where each object is displayed in a separate column).

### ClosedXML Integration
Generic.Excel is based on the ClosedXML library, which is a popular and powerful open-source library for working with Excel files in .NET.
By leveraging ClosedXML, Generic.Excel provides a reliable and efficient solution for generating XLSX files.
    
## Usage

 1- Clone or download the repository: To get started, clone or download the repository to your local machine.
 
 2- Open the solution file in Visual Studio 2019+: The solution file is located in the root directory of the project. Open this file in Visual Studio to start working with the project.
 
 3- Build the project in release mode
 
 4- Reference the class library dll in your application.
 
### Creating an Excel File

You can create an instance of the ExcelFile class using the static Create method:

```csharp
      using (ExcelFile excelFile = ExcelFile.Create())
      {
      }
```

### Adding Worksheets

To add a worksheet and export data, you can use the various AddSheet methods provided by the ExcelFile class:

```csharp
  List<MyModel> list = GetMyModelList();
  excelFile.AddSheetList(list, "Sheet1");
```

### Explicit properties to export

Adding a List of Objects with explicit properties to export:

```csharp
  List<MyModel> list = GetMyModelList();
  List<string> propertiesToExport = new List<string> { "Id", "Name", "Description" };
  excelFile.AddSheetList(list, "Sheet1", propertiesToExport);
```

### Explicit properties to export and display names

Adding a List of Objects with explicit properties to export and display names, explicit display names override decorator config.

If you want to export specific properties and define custom display names, use the AddSheetList method with propertiesToExport and propertiesDisplayName parameters:

```csharp
  List<MyModel> list = GetMyModelList();
  List<string> propertiesToExport = new List<string> { "Id", "Name", "Description" };
  List<string> propertiesDisplayName = new List<string> { "ID", "Name", "Description" };
  excelFile.AddSheetList(list, "Sheet1", propertiesToExport, propertiesDisplayName);
```

### Explicit properties to export and display names with dictionary

Alternatively, you can use a dictionary to map properties to their display names:

```csharp
  List<MyModel> list = GetMyModelList();
  Dictionary<string, string> propertiesNamesDictionary = new Dictionary<string, string>
  {
      { "Id", "ID" },
      { "Name", "Name" },
      { "Description", "Description" }
  };
  excelFile.AddSheetList(list, "Sheet1", propertiesNamesDictionary);
```

### Adding a List of Primitive Types

To export a list of primitive types, such as integers, use the AddSheetList method with a single property name:

```csharp
  List<int> numbers = GetNumbersList();
  excelFile.AddSheetList(numbers, "Sheet1", "Numbers");
```

### Adding a List of Strings

For a list of strings, you can export them using the AddSheet method:

```csharp
  List<string> strings = GetStringsList();
  excelFile.AddSheet(strings, "Sheet1", "Strings");
```

### Adding a Dictionary

To export a dictionary, use the AddSheet method with key and value property names:

```csharp
  Dictionary<string, int> dictionary = GetDictionary();
  excelFile.AddSheet(dictionary, "Sheet1", "Key", "Value");
```

### Adding a single object

To export a single object, use the AddSheetObj method:

```csharp
  MyModel singleObject = new Model{Name = "name", Description = "description"};
  excelFile.AddSheetObj<Model>(model, "Sheet1");
```

### Changing Export Direction

By default, Generic.Excel exports data in a vertical direction. If you want to change the export direction to horizontal:

```csharp
    List<MyModel> list = GetMyModelList();
    excelFile.AddSheetList(list, "Sheet1", horizontal: true);
```

### ExcelPropertyAttribute

Here's an example of using the ExcelPropertyAttribute::

```csharp
  public class MyModel
    {
        [ExcelProperty("ID", order: 1)]
        public int Id { get; set; }

        [ExcelProperty("Name", order: 3)]
        public string Name { get; set; }

        [ExcelProperty("Description", order: 4, ignore: true)]
        public string Description { get; set; }

        [ExcelProperty("Nested Property Name", order: 2, nestedProperty: "FullName")]
        public NestedModel NestedProperty { get; set; }

    }

    public class NestedModel
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public string FullName { get; set; }
    }
```

### Saving the Excel File

To save the generated Excel file to a specified location, use the SaveAs method:

```csharp
  excelFile.SaveAs("path/to/file.xlsx");
```

### Disposing the Excel File

Without the using {} declatation, it is important to release the resources used by the ExcelFile object when you are done with it. Call the Dispose method to free the resources:

```csharp
  excelFile.Dispose();
```

## RandomSharp Project

A lightweight random data generator for .NET.

## Support
If you are having problems, please let us know by [raising a new issue](https://github.com/fouad-elouad/Generic.Excel/issues/new/choose).

## License
This project is licensed with the [MIT license](LICENSE).